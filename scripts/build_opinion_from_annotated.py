# -*- coding: utf-8 -*-
"""从（人工调整后的）**批注版 docx** 回读批注 → 合并表类（明细表/测算表）意见 → 生成
「三级审核意见及回复记录」docx（v2 模板 / 索引号 G-18-3）。

为什么要有这个脚本（2026-08-25 用户定的新流程）：
  交付顺序改为 **先批注版 + 交叉验证 → 停下来请人工过目调整批注 → 再出审核意见**。
  审核意见的"事实来源"因此**不再是模型记忆里的那份意见清单**，而是**当下 docx 里真实存在的批注**。
  人工在 Word 里删掉/改写/新增了哪几条，本脚本回读到的就是哪几条，
  序号在生成时按 1..N 现场重排 —— **删掉任意一条，序号仍然连续**。

用法：
    python build_opinion_from_annotated.py 配置.json            # 出成品
    python build_opinion_from_annotated.py 配置.json --preview   # 只打印逐条清单，不产文件
    python build_opinion_from_annotated.py --dump 批注版.docx    # 只看某份 docx 里现有批注

配置 JSON 键：
    PROJECT_NAME / PROJECT_NO / REVIEWER / REVIEW_DATE /
    SECOND_REVIEWER / REPLIER / REPLY_DATE / OUTPUT      —— 同 build_opinion.py
    SOURCES     : [ {LABEL, DOCX, AUTHORS?, POSITION_PREFIX?, POSITION_FALLBACK?}, ... ]
                  按顺序即为大类顺序（资产评估报告 → 评估说明 → …）；
                  AUTHORS 为批注人姓名白名单，**强烈建议填**（滤掉承做人残留在源文档里的旧批注）。
    EXTRA_GROUPS: [ [类别提示, [[位置, 原文, 审核意见], ...]], ... ]
                  Excel 类（评估明细表/评估测算表）等无 Word 批注的意见，手工给，排在 SOURCES 之后。
    DROP        : 可选。整数 = 按上一次 preview 的序号删该条；字符串 = 删意见正文含该子串的条。
                  删完**自动重排序号**，无需手改任何编号。
    MAX_QUOTE   : 可选，原文截断字数，默认 60（超出加"……"）；填 0 = 不截断。
    REPLY_MODE  : 可选 skip(默认)/merge/separate —— 人工在 Word 里"答复"形成的子批注怎么处理。
    INCLUDE_DONE: 可选，默认 false —— Word 里标为"已解决"的批注默认**不进**审核意见。
    DUMP_JSON   : 可选，把生成的 build_opinion 配置落盘（便于人工微调后直接重跑 build_opinion.py）。
"""
import sys, os, re, json, zipfile, subprocess
import xml.etree.ElementTree as ET

try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
BUILD_OPINION = os.path.join(HERE, "build_opinion.py")

W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
def q(tag, ns=W): return '{%s}%s' % (ns, tag)

# 标题识别：先看样式，再看"第X章/一、/（一）/1.1"这类中文报告惯用写法（多数报告不套 Heading 样式）
# 三级：第X章·一、 / （一） / 1.、1.1
RE_LVL1 = re.compile(r'^\s*(?:第[一二三四五六七八九十百]+\s*[章部编篇]|[一二三四五六七八九十]+\s*、)')
RE_LVL2 = re.compile(r'^\s*[（(][一二三四五六七八九十]+[）)]')
RE_LVL3 = re.compile(r'^\s*\d+(?:\.\d+)*\s*[、．.]')
RE_LVLD = re.compile(r'^\s*[（(]\d+[）)]')
RE_CJK  = re.compile(r'[一-鿿]')


def _ptext(p):
    return ''.join(t.text or '' for t in p.iter(q('t'))).strip()


def _style_level(p):
    """返回样式判定的标题级别（1/2），不是标题返回 0。"""
    ppr = p.find(q('pPr'))
    if ppr is None: return 0
    st = ppr.find(q('pStyle'))
    val = (st.get(q('val')) if st is not None else '') or ''
    if 'Heading' in val or '标题' in val:
        m = re.search(r'(\d+)', val)
        return 1 if (m and m.group(1) == '1') else 2
    lvl = ppr.find(q('outlineLvl'))
    if lvl is not None:
        try: return 1 if int(lvl.get(q('val'))) == 0 else 2
        except Exception: return 2
    return 0


def _heading_type(p, txt, in_table=False):
    """返回标题的**编号类型**（不是固定层级）：
        A=第X章／一、二、   B=（一）（二）   C=1.／1.1   D=（1）（2）   S=仅靠样式判定
    不是标题返回 None。用类型而非固定层级，是因为同一份报告里两种体例的嵌套顺序可能相反：
    结果报告是「三、＞（四）＞2.」，测算过程却是「一、＞1.＞（2）」——固定层级必然串位。
    """
    lv = _heading_level(p, txt, in_table)
    if not lv: return None
    if RE_LVL1.match(txt): return 'A'
    if RE_LVL2.match(txt): return 'B'
    if RE_LVL3.match(txt): return 'C'
    if RE_LVLD.match(txt): return 'D'
    return 'S'


def _push_heading(stack, t, txt):
    """按编号类型维护标题栈：同类型出现即替换并截断其下各级；新类型则入栈。"""
    if t == 'A':
        stack[:] = [('A', txt)]; return
    for i, (k, _) in enumerate(stack):
        if k == t:
            stack[:] = stack[:i] + [(t, txt)]; return
    if len(stack) < 4: stack.append((t, txt))
    else: stack[-1] = (t, txt)


def _heading_level(p, txt, in_table=False):
    """综合样式与中文序号写法判定标题级别；返回 0 表示不是标题。
    ⚠️两条硬门槛（来自实测教训）：
      ①**表格单元格内的段落一律不当标题** —— 报告里的表头"（一）建筑物""（二）土地使用权"
        和取费表里的"2.92%"都会被序号规则命中，冒充章节名污染「位置」字段；
      ②标题须含≥2个汉字且不超过40字，挡掉"2.92%""1.68"这类纯数字/百分比单元格。
    """
    if in_table: return 0
    if not txt or len(txt) > 40: return 0
    if len(RE_CJK.findall(txt)) < 2: return 0
    # 中文报告的"第X章/一、/（一）/1."层级最可靠，优先于样式名
    # （很多报告把各级标题一律套 Heading1，只看样式会把"一、"误判成一级、把上级章名冲掉）
    if RE_LVL1.match(txt): return 1
    if RE_LVL2.match(txt): return 2
    if RE_LVL3.match(txt): return 3
    if RE_LVLD.match(txt): return 4
    return _style_level(p)


def _read_comment_bodies(z):
    """word/comments.xml → {id: {author, text, paraids:[...]}}"""
    out = {}
    if 'word/comments.xml' not in z.namelist(): return out
    root = ET.fromstring(z.read('word/comments.xml'))
    for c in root.findall(q('comment')):
        cid = c.get(q('id'))
        paras = []
        for p in c.iter(q('p')):
            pid = p.get(q('paraId', W14))
            if pid: paras.append(pid.lower())
        # 批注正文可能多段，段间用换行连起来（审核意见正文含 \n 会被 build_opinion 拆成多段）
        lines = [_ptext(p) for p in c.iter(q('p'))]
        txt = '\n'.join(x for x in lines if x).strip()
        if not txt:
            txt = ''.join(t.text or '' for t in c.iter(q('t'))).strip()
        out[cid] = {'author': (c.get(q('author')) or '').strip(),
                    'initials': (c.get(q('initials')) or '').strip(),
                    'date': (c.get(q('date')) or ''),
                    'text': txt, 'paraids': paras}
    return out


def _read_comments_ex(z, bodies):
    """word/commentsExtended.xml → 每条批注的 done / 是否为回复（paraIdParent）。"""
    info = {cid: {'done': False, 'parent': None} for cid in bodies}
    if 'word/commentsExtended.xml' not in z.namelist(): return info
    pid2cid = {}
    for cid, b in bodies.items():
        for pid in b['paraids']: pid2cid[pid] = cid
    try:
        root = ET.fromstring(z.read('word/commentsExtended.xml'))
    except Exception:
        return info
    for ex in root.iter(q('commentEx', W15)):
        pid = (ex.get(q('paraId', W15)) or '').lower()
        cid = pid2cid.get(pid)
        if cid is None: continue
        done = (ex.get(q('done', W15)) or '0') in ('1', 'true', 'True')
        parent_pid = (ex.get(q('paraIdParent', W15)) or '').lower()
        info[cid]['done'] = info[cid]['done'] or done
        if parent_pid: info[cid]['parent'] = pid2cid.get(parent_pid)
    return info


def read_comments(path):
    """按**文档顺序**回读一份 docx 里现存的全部批注。
    返回 [{id, author, text, quote, para, heading, done, parent}]，quote=批注锚定范围内的原文。
    """
    z = zipfile.ZipFile(path)
    bodies = _read_comment_bodies(z)
    exinfo = _read_comments_ex(z, bodies)
    doc = ET.fromstring(z.read('word/document.xml'))
    body = doc.find(q('body'))
    if body is None: body = doc

    order, meta = [], {}
    ranges = {}          # id -> 锚定范围内收集到的文本
    active = set()
    stack = []                       # [(编号类型, 标题文字), ...]，见 _push_heading
    state = {'para': ''}

    def heading():
        return _cut(' '.join(t for _, t in stack), 60)

    def note(cid):
        if cid in meta: return
        order.append(cid)
        meta[cid] = {'heading': heading(), 'para': state['para']}

    def walk(el, in_table):
        """按文档顺序深度遍历，同时带着"是否在表格内"的上下文。"""
        for child in list(el):
            tag = child.tag
            if tag == q('tbl'):
                walk(child, True); continue
            if tag == q('p'):
                txt = _ptext(child)
                state['para'] = txt
                t = _heading_type(child, txt, in_table)
                if t: _push_heading(stack, t, txt)
            elif tag == q('commentRangeStart'):
                cid = child.get(q('id')); active.add(cid); ranges.setdefault(cid, []); note(cid)
            elif tag == q('commentRangeEnd'):
                active.discard(child.get(q('id')))
            elif tag == q('commentReference'):
                note(child.get(q('id')))   # 只有引用点、没有 range（Word 里点在插入点上的批注）
            elif tag == q('t') and active:
                for cid in active: ranges[cid].append(child.text or '')
            if len(child):
                walk(child, in_table)

    walk(body, False)

    # comments.xml 里有、正文里已找不到锚点的（人工删了正文却留了批注体），按 id 顺序补在末尾
    for cid in sorted(bodies, key=lambda x: int(x) if str(x).isdigit() else 0):
        if cid not in meta:
            order.append(cid); meta[cid] = {'heading': '', 'para': ''}

    out = []
    for cid in order:
        b = bodies.get(cid)
        if b is None: continue            # 有锚点没批注体 = 残留标记，跳过
        quote = ''.join(ranges.get(cid, [])).strip() or meta[cid]['para']
        out.append({'id': cid, 'author': b['author'], 'text': b['text'],
                    'quote': quote.strip(), 'para': meta[cid]['para'],
                    'heading': meta[cid]['heading'],
                    'done': exinfo.get(cid, {}).get('done', False),
                    'parent': exinfo.get(cid, {}).get('parent')})
    return out


def _cut(s, n):
    s = re.sub(r'\s+', ' ', s or '').strip()
    if n and len(s) > n: return s[:n] + '……'
    return s


def collect(cfg):
    """按 SOURCES + EXTRA_GROUPS 汇总成 build_opinion 的 GROUPS，并返回统计信息。"""
    max_quote  = cfg.get('MAX_QUOTE', 60)
    reply_mode = (cfg.get('REPLY_MODE') or 'skip').lower()
    inc_done   = bool(cfg.get('INCLUDE_DONE', False))
    notes, groups = [], []

    for src in cfg.get('SOURCES', []):
        label = src.get('LABEL') or ''
        path  = src.get('DOCX')
        if not path or not os.path.exists(path):
            notes.append('[WARN] 批注版不存在，已跳过：%s' % path); continue
        authors = src.get('AUTHORS') or []
        prefix  = src.get('POSITION_PREFIX') or ''
        fallback = src.get('POSITION_FALLBACK') or ''
        cms = read_comments(path)
        kept, dropped_author, dropped_done, replies = [], 0, 0, 0
        by_id = {c['id']: c for c in cms}
        for c in cms:
            if authors and c['author'] not in authors:
                dropped_author += 1; continue
            if c['parent'] is not None:
                replies += 1
                if reply_mode == 'skip': continue
                if reply_mode == 'merge':
                    p = by_id.get(c['parent'])
                    if p is not None:
                        p['text'] = (p['text'] + '\n' + c['text']).strip(); continue
            if c['done'] and not inc_done:
                dropped_done += 1; continue
            kept.append(c)
        items = []
        for c in kept:
            pos = ' '.join(x for x in (prefix, c['heading']) if x).strip() or fallback
            items.append([pos, _cut(c['quote'], max_quote), c['text'].strip()])
        notes.append('[读] %s ← %s：现存批注 %d 条，取用 %d 条'
                     % ((label or '(无类别提示)').rstrip('：:'), os.path.basename(path), len(cms), len(items)))
        if dropped_author:
            notes.append('      · 按 AUTHORS 过滤掉 %d 条（非本次批注人，如承做人旧批注）' % dropped_author)
        if replies:
            notes.append('      · 回复型子批注 %d 条，REPLY_MODE=%s' % (replies, reply_mode))
        if dropped_done:
            notes.append('      · Word 中标为"已解决" %d 条，未计入（如需计入设 INCLUDE_DONE=true）' % dropped_done)
        if items: groups.append([label, items])

    for g in cfg.get('EXTRA_GROUPS', []):
        label, items = g[0], [list(x) for x in g[1]]
        items = [[_cut(a, 0), _cut(b, max_quote), (c or '').strip()] for a, b, c in items]
        notes.append('[表] %s：手工意见 %d 条' % (label.rstrip('：:'), len(items)))
        if items: groups.append([label, items])
    return groups, notes


def apply_drop(groups, drop):
    """DROP：整数按序号删、字符串按意见正文子串删；删完序号由 build_opinion 现场 1..N 重排。"""
    if not drop: return groups, []
    ints = set(int(d) for d in drop if isinstance(d, int) or str(d).isdigit())
    subs = [str(d) for d in drop if not (isinstance(d, int) or str(d).isdigit())]
    out, log, n = [], [], 0
    for label, items in groups:
        keep = []
        for it in items:
            n += 1
            if n in ints or any(s and s in it[2] for s in subs):
                log.append('[DROP] 原序号 %d：%s' % (n, _cut(it[2], 30)))
            else:
                keep.append(it)
        if keep: out.append([label, keep])
    return out, log


def preview(groups):
    n = 0
    for label, items in groups:
        print('\n== %s ==' % label)
        for pos, quote, op in items:
            n += 1
            print('%2d、位置：%s' % (n, pos))
            print('    原文：%s' % _cut(quote, 40))
            print('    意见：%s' % op.replace('\n', ' / '))
    print('\n合计 %d 条（序号 1..%d 连续）' % (n, n))
    return n


def main():
    args = [a for a in sys.argv[1:]]
    if '--dump' in args:
        i = args.index('--dump')
        for c in read_comments(args[i + 1]):
            print(json.dumps(c, ensure_ascii=False))
        return 0
    if not args:
        print('用法：python build_opinion_from_annotated.py 配置.json [--preview]'); return 2
    cfg = json.load(open(args[0], encoding='utf-8'))

    groups, notes = collect(cfg)
    for x in notes: print(x)
    groups, droplog = apply_drop(groups, cfg.get('DROP'))
    for x in droplog: print(x)
    n = preview(groups)
    if n == 0:
        print('[FAIL] 没有任何意见可写入，请检查 SOURCES 路径/AUTHORS 是否过滤过头'); return 1
    if '--preview' in args:
        print('\n（--preview 模式，未生成文件）'); return 0

    sub = {k: cfg.get(k, '') for k in ('PROJECT_NAME', 'PROJECT_NO', 'REVIEWER', 'REVIEW_DATE',
                                       'SECOND_REVIEWER', 'REPLIER', 'REPLY_DATE', 'OUTPUT')}
    sub['GROUPS'] = groups
    dump = cfg.get('DUMP_JSON') or os.path.join(
        os.path.dirname(os.path.abspath(cfg.get('OUTPUT') or args[0])), '_审核意见_合并配置.json')
    with open(dump, 'w', encoding='utf-8') as f:
        json.dump(sub, f, ensure_ascii=False, indent=1)
    print('\n[合并配置] %s' % dump)
    print('  （如需再微调文字，可直接改这份 JSON 后跑：python build_opinion.py "%s"）' % dump)
    # 捕获子进程输出再转印，确保 build_opinion 的强制自检结论在任何终端/管道下都能看到
    r = subprocess.run([sys.executable, BUILD_OPINION, dump], capture_output=True)
    for stream in (r.stdout, r.stderr):
        if stream:
            print(stream.decode('utf-8', 'replace').rstrip())
    if r.returncode != 0:
        print('[FAIL] build_opinion.py 自检未通过，未产出成品')
    return r.returncode


if __name__ == '__main__':
    sys.exit(main())
