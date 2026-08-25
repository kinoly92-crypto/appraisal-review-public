# -*- coding: utf-8 -*-
"""生成「三级审核意见及回复记录」docx（模板版本 v2 / 索引号 G-18-3）。

工程化要点（保证换电脑、换大模型输出仍高度一致）：
  1. 成品 = 克隆 templates/opinion-template.docx，**除 word/document.xml 外逐字节原样保留**
     （样式表/字体表/主题/页面设置全部继承模板，不由脚本或模型自由发挥）。
  2. 所有排版 XML 片段是本文件里的常量，模型只提供「内容」（位置/原文/审核意见），不碰格式。
  3. 模板三张表按「列数签名」定位（信息表2列 / 正文表1列 / 签字表3列），不依赖示例文字，
     模板即使被人改了示例内容也能正确落位。
  4. 收尾做强制自检：XML 合法性、占位符已替换、序号 1..N 连续、信息表已写入；
     任一不过直接非零退出并报错，**绝不产出"看起来对"的错文档**。

用法（跨平台首选 JSON，免中文乱码）：
    python build_opinion.py 配置.json
JSON 键名 = 下面【按项目填写】区的大写变量名。

模板结构（v2，2026-08 起启用）：
    标题「对《X》三级审核意见及回复记录」→ 右对齐「索引号：G-18-3」
    → 信息表(6行2列: 项目编号/二级复核人/审核人/审核日期/回复人/回复日期)
    → 分隔线 → 正文表(1×1 无框，所有意见条目装在里面) → 分隔线
    → 签字表(3×3: 修改人签字·签字日期 / 审核人验证意见 / 审核人签字·签字日期)
"""
import sys, zipfile, re, os, json, datetime
import xml.dom.minidom as M
try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "templates", "opinion-template.docx")

# ===================== 【按项目填写】 =====================
PROJECT_NAME = "<完整项目名，不含书名号>"   # 写入标题《》内
PROJECT_NO   = "<项目编号>"
REVIEWER     = "<审核人姓名>"               # 本次问询所得，勿写死
REVIEW_DATE  = "2026年X月X日"
SECOND_REVIEWER = ""                        # 二级复核人，通常留空待填
REPLIER      = ""                           # 回复人，留空
REPLY_DATE   = ""                           # 回复日期，留空
OUTPUT       = r"<输出docx绝对路径，建议放项目文件夹>"
# 分组：每个大类 = (类别提示, [ (位置, 原文, 审核意见正文), ... ] )
# 顺序固定：资产评估报告 → 评估说明 → 评估明细表 → 评估测算表（房地产/土地：报告 → 测算表）
# 审核意见正文可含 "\n"，将拆成多个缩进段。位置/原文允许留空字符串（该行不输出）。
GROUPS = [
 ("资产评估报告：", [
    ("<位置>", "<原文>", "<审核意见：只描述问题，不写依据条款、不写整改建议>"),
 ]),
 ("评估测算表：", [
    ("<位置>", "<原文>", "<审核意见>"),
 ]),
]
# 兼容旧配置：若只给了 TITLE（形如 "对《X》审核意见"），自动从中提取项目名
TITLE = ""
# =========================================================

if len(sys.argv) > 1:
    _c = json.load(open(sys.argv[1], encoding='utf-8'))
    PROJECT_NAME = _c.get('PROJECT_NAME', PROJECT_NAME)
    TITLE        = _c.get('TITLE', TITLE)
    PROJECT_NO   = _c.get('PROJECT_NO', PROJECT_NO)
    REVIEWER     = _c.get('REVIEWER', REVIEWER)
    REVIEW_DATE  = _c.get('REVIEW_DATE', REVIEW_DATE)
    SECOND_REVIEWER = _c.get('SECOND_REVIEWER', SECOND_REVIEWER)
    REPLIER      = _c.get('REPLIER', REPLIER)
    REPLY_DATE   = _c.get('REPLY_DATE', REPLY_DATE)
    OUTPUT       = _c.get('OUTPUT', OUTPUT)
    GROUPS       = _c.get('GROUPS', GROUPS)

# 旧配置兼容：TITLE="对《X》审核意见" → PROJECT_NAME=X
if TITLE and (not PROJECT_NAME or PROJECT_NAME.startswith('<')):
    _m = re.search(r'《(.+?)》', TITLE)
    if _m: PROJECT_NAME = _m.group(1)
PROJECT_NAME = (PROJECT_NAME or '').strip().strip('《》')
if not REVIEW_DATE or REVIEW_DATE.startswith('2026年X月'):
    _t = datetime.date.today(); REVIEW_DATE = "%d年%d月%d日" % (_t.year, _t.month, _t.day)

# 标题两段（「对《项目名》」／「三级审核意见及回复记录」）由**模板自带**，脚本只替换占位符，勿再自行拆分
TITLE_LINE2 = "三级审核意见及回复记录"
PLACEHOLDER = "[项目名称]"
PH_NO, PH_REVIEWER, PH_DATE = "[项目编号]", "[审核人]", "[审核日期]"   # 信息表占位符，由 fill_info_table 覆盖
INDEX_NO    = "索引号：G-18-3"          # 三级审核专用表单号，勿改

# ---------------- XML 片段常量（排版唯一来源，勿在别处另写） ----------------
def esc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

RFONTS = ('<w:rFonts w:hint="eastAsia" w:ascii="Times New Roman" w:hAnsi="Times New Roman"'
          ' w:eastAsia="仿宋_GB2312" w:cs="Times New Roman"/>')

# ---------------- 序号＝Word 自动编号（2026-08-25 用户要求） ----------------
# 为什么不用字面文本「1、」：用户会直接在 Word 里删掉整条意见，字面序号删完就断号。
# 挂真正的编号列表后，Word 自己重排，删任意一条后面自动顺延。
# v2 模板不带 numbering.xml，故本脚本从零生成该部件并挂进包里（Content_Types + rels）。
NUM_ID, ABSTRACT_ID = 20, 20        # 模板无编号列表，取一个不会冲突的编号
NUMBERING_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="%d">'
    '<w:multiLevelType w:val="singleLevel"/>'
    '<w:lvl w:ilvl="0">'
    '<w:start w:val="1"/>'
    '<w:numFmt w:val="decimal"/>'
    '<w:suff w:val="nothing"/>'          # 序号后不加制表符/空格，与原字面「1、」观感一致
    '<w:lvlText w:val="%%1、"/>'
    '<w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="0" w:leftChars="0" w:firstLine="0" w:firstLineChars="0"/></w:pPr>'
    '<w:rPr>%s<w:b w:val="0"/><w:bCs w:val="0"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
    '</w:lvl></w:abstractNum>'
    '<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'
    '</w:numbering>') % (ABSTRACT_ID, RFONTS, NUM_ID, ABSTRACT_ID)
CT_NUMBERING = ('<Override PartName="/word/numbering.xml" ContentType="application/vnd.'
                'openxmlformats-officedocument.wordprocessingml.numbering+xml"/>')
REL_NUMBERING = ('<Relationship Id="rIdNum%d" Type="http://schemas.openxmlformats.org/'
                 'officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' % NUM_ID)

def rpr(bold=False, sz=24):
    b = '<w:b/><w:bCs/>' if bold else '<w:b w:val="0"/><w:bCs w:val="0"/>'
    return '<w:rPr>%s%s<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (RFONTS, b, sz, sz)

def run(text, bold=False, sz=24):
    if text == '': return ''
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr(bold, sz), esc(text))

def para(runs_xml, bdr=False, indent=False, bold_mark=False, num=False):
    """一个段落。bdr=段落下边框(条目分隔线)；indent=首行缩进2字符；
    num=True 挂 Word **自动编号**（序号段专用，见 NUM_ID）。
    pPr 子元素顺序须严格符合 CT_PPr schema 序列，勿随意调换。"""
    p = ['<w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/>'
         '<w:widowControl w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>'
         % (NUM_ID if num else 0)]
    if bdr:
        p.append('<w:pBdr><w:bottom w:val="single" w:color="auto" w:sz="4" w:space="0"/></w:pBdr>')
    p.append('<w:kinsoku/><w:wordWrap/><w:overflowPunct/><w:topLinePunct w:val="0"/>'
             '<w:autoSpaceDE/><w:autoSpaceDN/><w:bidi w:val="0"/><w:adjustRightInd/><w:snapToGrid/>'
             '<w:spacing w:line="360" w:lineRule="auto"/>')
    p.append('<w:ind w:left="0" w:leftChars="0" w:firstLine="480" w:firstLineChars="200"/>' if indent
             else '<w:ind w:left="0" w:leftChars="0" w:firstLine="0" w:firstLineChars="0"/>')
    p.append('<w:jc w:val="both"/><w:textAlignment w:val="auto"/>')
    p.append(rpr(bold_mark))
    p.append('</w:pPr>')
    return '<w:p>' + ''.join(p) + runs_xml + '</w:p>'

def cat_header(text):
    """大类提示行（加粗），如"资产评估报告："。"""
    return para(run(text, bold=True), bold_mark=True)

def item_block(n, loc, orig, opinion, last=False):
    """一条意见块（尾部三段照模板：回复：[下边框] + 空段[下边框] + 空段）：
       N、／位置：／原文：／审核意见：／正文／回复：[下边框]／空段[下边框]／空段
    位置、原文为空则该行不输出；审核意见正文含 \n 时拆成多个缩进段。
    last=True（全篇最后一条）省去最后那个不带边框的空段。"""
    out = [para('', num=True)]        # 序号段：空文本，「N、」由 Word 自动编号生成（删条自动重排）
    if loc:  out.append(para(run('位置：', True) + run(loc)))
    if orig: out.append(para(run('原文：', True) + run(orig)))
    out.append(para(run('审核意见：', True)))
    for seg in [s for s in str(opinion).split('\n') if s.strip()] or ['']:
        out.append(para(run(seg), indent=True))
    out.append(para(run('回复：', True), bdr=True))     # 回复：+下边框
    out.append(para('', bdr=True))                              # 空段+下边框
    if not last:
        out.append(para(''))                                    # 空段（不带边框）
    return ''.join(out)

def find_para(doc, needle):
    """返回包含 needle 的第一个 <w:p>…</w:p> 匹配（段落不嵌套，非贪婪即可）。"""
    for m in re.finditer(r'<w:p(?:\s[^>]*)?>.*?</w:p>', doc, re.S):
        if needle in m.group(0): return m
    return None

# ---------------- 模板结构定位（按列数签名，不依赖示例文字） ----------------
def find_tables(doc):
    """返回顶层 <w:tbl> 的 (start, end) 列表。"""
    spans, depth, start = [], 0, None
    for m in re.finditer(r'<w:tbl>|</w:tbl>', doc):
        if m.group(0) == '<w:tbl>':
            if depth == 0: start = m.start()
            depth += 1
        else:
            depth -= 1
            if depth == 0: spans.append((start, m.end()))
    return spans

def pick_table(doc, spans, ncols):
    hits = [s for s in spans if doc[s[0]:s[1]].count('<w:gridCol') == ncols]
    if len(hits) != 1:
        sys.exit('[FAIL] 模板结构异常：%d 列的表格找到 %d 张（应为 1 张）。模板可能被改动，请核对 templates/opinion-template.docx' % (ncols, len(hits)))
    return hits[0]

def strip_runs(xml):
    xml = re.sub(r'<w:r>.*?</w:r>', '', xml, flags=re.S)
    return re.sub(r'<w:r\s[^>]*>.*?</w:r>', '', xml, flags=re.S)

def fill_info_table(tbl_xml, pairs):
    """信息表：把每个 label 所在单元格的**下一个**单元格内容置为 value。
    值单元格若已带占位符 run（如 [项目编号]），沿用该 run 的 rPr 后再换文字，
    这样模板改字体字号时成品自动跟随，不必改脚本。"""
    for label, value in pairs:
        tcs = list(re.finditer(r'<w:tc>.*?</w:tc>', tbl_xml, re.S))
        idx = None
        for i, m in enumerate(tcs):
            if label in re.sub(r'<[^>]+>', '', m.group(0)):
                idx = i; break
        if idx is None or idx + 1 >= len(tcs):
            sys.exit('[FAIL] 模板信息表缺少标签：%s' % label)
        vc = tcs[idx + 1]; raw = vc.group(0)
        m = re.search(r'<w:r>\s*(<w:rPr>.*?</w:rPr>)', raw, re.S)      # 占位符 run 的格式
        vrpr = m.group(1) if m else rpr(False)
        xml = strip_runs(raw)
        if value:
            j = xml.index('</w:p>')
            xml = (xml[:j] + '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (vrpr, esc(value))
                   + xml[j:])
        tbl_xml = tbl_xml[:vc.start()] + xml + tbl_xml[vc.end():]
    return tbl_xml

# ---------------- 组装 ----------------
if not os.path.exists(TEMPLATE):
    sys.exit('[FAIL] 找不到模板：%s' % TEMPLATE)
zin = zipfile.ZipFile(TEMPLATE)
doc = zin.read('word/document.xml').decode('utf-8')

# 1) 标题：模板已自带两段（「对《[项目名称]》」/「三级审核意见及回复记录」），此处只替换占位符，
#    **不要再自行拆段**，否则会多出一行标题。
tp = find_para(doc, PLACEHOLDER)
if tp is None:
    sys.exit('[FAIL] 模板标题里找不到占位符 %s，请用 templates/opinion-template.docx 原件重跑。' % PLACEHOLDER)
doc = doc[:tp.start()] + tp.group(0).replace(PLACEHOLDER, esc(PROJECT_NAME)) + doc[tp.end():]
if INDEX_NO.replace('：', '') not in re.sub(r'<[^>]+>', '', doc).replace('：', ''):
    print('[WARN] 模板中未检出索引号 %s，请核对模板版本。' % INDEX_NO)

spans = find_tables(doc)
if len(spans) != 3:
    sys.exit('[FAIL] 模板应含 3 张表（信息表/正文表/签字表），实际 %d 张。' % len(spans))
info_span, body_span, sign_span = pick_table(doc, spans, 2), pick_table(doc, spans, 1), pick_table(doc, spans, 3)

# 2) 正文表：重建单元格内容
body_xml = doc[body_span[0]:body_span[1]]
m = re.search(r'(<w:tc>)(<w:tcPr>.*?</w:tcPr>)(.*)(</w:tc>)', body_xml, re.S)
if not m:
    sys.exit('[FAIL] 正文表单元格结构异常。')
_groups = [(hd, lst) for hd, lst in GROUPS if lst]      # 无意见的大类不输出
_total = sum(len(lst) for _, lst in _groups)
if _total == 0:
    sys.exit('[FAIL] GROUPS 里没有任何意见条目。')
n = 0; content = []
for hd, lst in _groups:
    content.append(cat_header(hd))
    for it in lst:
        loc, orig, opinion = (list(it) + ['', '', ''])[:3]
        n += 1
        content.append(item_block(n, loc, orig, opinion, last=(n == _total)))
new_body = body_xml[:m.start()] + m.group(1) + m.group(2) + ''.join(content) + m.group(4) + body_xml[m.end():]
doc = doc[:body_span[0]] + new_body + doc[body_span[1]:]

# 3) 信息表（正文表改动会移动偏移，故重新定位）
spans = find_tables(doc); info_span = pick_table(doc, spans, 2)
info_xml = fill_info_table(doc[info_span[0]:info_span[1]], [
    ('项目编号：', PROJECT_NO),
    ('二级复核人：', SECOND_REVIEWER),
    ('审核人：', REVIEWER),
    ('审核日期：', REVIEW_DATE),
    ('回复人：', REPLIER),
    ('回复日期：', REPLY_DATE),
])
doc = doc[:info_span[0]] + info_xml + doc[info_span[1]:]

# ---------------- 强制自检（不过就报错退出，不产出错文档） ----------------
errs = []
try: M.parseString(doc.encode('utf-8'))
except Exception as e: errs.append('document.xml 不合法：%s' % e)
plain = re.sub(r'<[^>]+>', '', doc)
plain = plain.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')   # 反转义后再比对
for ph in (PLACEHOLDER, PH_NO, PH_REVIEWER, PH_DATE):
    if ph in plain: errs.append('占位符 %s 未替换' % ph)
if PROJECT_NAME and PROJECT_NAME[:12] not in plain: errs.append('标题未写入项目名')
_t2 = find_para(doc, TITLE_LINE2)          # 承载第二行的那一段里不应再有「对《」
if _t2 is None or '对《' in _t2.group(0): errs.append('标题两行被合并成一段（应分两段）')
if plain.count(TITLE_LINE2) != 1: errs.append('标题第二行出现 %d 次（应为 1 次）' % plain.count(TITLE_LINE2))
if doc.count('<w:t xml:space="preserve">回复：</w:t>') != n:
    errs.append('「回复：」段数 %d ≠ 意见条数 %d' % (doc.count('<w:t xml:space="preserve">回复：</w:t>'), n))
_bs = find_tables(doc)
_bt = doc[pick_table(doc, _bs, 1)[0]:pick_table(doc, _bs, 1)[1]]
if _bt.count('<w:pBdr>') != 2 * n:
    errs.append('正文表内下边框段 %d 个（每条应 2 个，共应 %d 个）' % (_bt.count('<w:pBdr>'), 2 * n))
_numbered = doc.count('<w:numId w:val="%d"/>' % NUM_ID)
if _numbered != n:
    errs.append('挂自动编号的序号段 %d 个 ≠ 意见条数 %d' % (_numbered, n))
if re.search(r'<w:t[^>]*>\d+、</w:t>', doc):
    errs.append('正文中仍存在字面序号（应全部改为 Word 自动编号）')
for lb, v in (('项目编号：', PROJECT_NO), ('审核人：', REVIEWER), ('审核日期：', REVIEW_DATE)):
    if v and (lb + v) not in plain.replace('\n', ''): errs.append('信息表未写入 %s%s' % (lb, v))
if len(find_tables(doc)) != 3: errs.append('成品表格数不为 3')
if errs:
    sys.exit('[FAIL] 自检未通过：\n  - ' + '\n  - '.join(errs))

# ---------------- 打包：把 numbering.xml 挂进包（模板不带该部件） ----------------
_ct = zin.read('[Content_Types].xml').decode('utf-8')
if 'numbering+xml' not in _ct:
    _ct = _ct.replace('</Types>', CT_NUMBERING + '</Types>')
_rels = zin.read('word/_rels/document.xml.rels').decode('utf-8')
if 'relationships/numbering' not in _rels:
    _rels = _rels.replace('</Relationships>', REL_NUMBERING + '</Relationships>')
try:
    M.parseString(NUMBERING_XML.encode('utf-8')); M.parseString(_ct.encode('utf-8')); M.parseString(_rels.encode('utf-8'))
except Exception as e:
    sys.exit('[FAIL] 编号部件 XML 不合法：%s' % e)

parts = []
_has_numbering = False
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename == 'word/document.xml': d = doc.encode('utf-8')
    elif it.filename == '[Content_Types].xml': d = _ct.encode('utf-8')
    elif it.filename == 'word/_rels/document.xml.rels': d = _rels.encode('utf-8')
    elif it.filename == 'word/numbering.xml':
        d = NUMBERING_XML.encode('utf-8'); _has_numbering = True
    parts.append((it, d))
if not _has_numbering:
    parts.append(('word/numbering.xml', NUMBERING_XML.encode('utf-8')))
target = OUTPUT
try:
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, d in parts: z.writestr(it, d)
except PermissionError:
    target = OUTPUT[:-5] + '（已修订）.docx'
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, d in parts: z.writestr(it, d)
    print('[锁定→另存]', target, '（原文件可能在 Word 中打开）')
print('[OK]', target)
print('  标题第1行：对《%s》' % PROJECT_NAME)
print('  标题第2行：%s' % TITLE_LINE2)
print('  %s | 项目编号 %s | 审核人 %s | 审核日期 %s' % (INDEX_NO, PROJECT_NO, REVIEWER, REVIEW_DATE))
print('  意见 %d 条，分 %d 个大类：%s' % (n, len(_groups), '、'.join(g[0].rstrip('：') for g in _groups)))
print('  自检全部通过（XML合法 / 4处占位符已替 / 标题分两行且不重复 / %d 条挂 Word 自动编号且无字面序号 / '
      '回复段 %d 个 / 下边框段 %d 个 / 信息表已写入 / 三表完整）' % (n, n, 2 * n))
print('  ⭐序号为 Word 自动编号：在 Word 里直接删除整条意见，后续序号自动重排，不会断号。')
