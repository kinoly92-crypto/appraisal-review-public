# -*- coding: utf-8 -*-
"""生成「上轮审核意见回复落实情况核查表」docx（v2 模板）。

与 build_opinion.py 同源：克隆 templates/opinion-template.docx，除 word/document.xml
外逐字节保留；三张表按**列数签名**(2/1/3)定位；收尾强制自检。
差别只在正文表内的块结构：每条是「上轮意见／承做人回复／落实核查／核查结论」四段。

用法：python build_followup.py 配置.json
键：PROJECT_NAME, PROJECT_NO, REVIEWER, REVIEW_DATE, OUTPUT, GROUPS
    可选 DOC_TITLE(标题第2行，默认"审核意见回复落实情况核查表")
        LABEL(上轮意见称谓，默认"上轮意见")、INTRO(表前说明段)
        SECOND_REVIEWER / REPLIER / REPLY_DATE
    GROUPS = [[大类提示, [[原意见号, 原意见, 承做人回复, 落实核查, 核查结论], …]], …]
"""
import sys, zipfile, re, os, json, datetime
import xml.dom.minidom as M
try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "templates", "opinion-template.docx")
if len(sys.argv) < 2:
    sys.exit('用法：python build_followup.py 配置.json')
_c = json.load(open(sys.argv[1], encoding='utf-8'))

PROJECT_NAME = (_c.get('PROJECT_NAME') or '').strip().strip('《》')
PROJECT_NO   = _c.get('PROJECT_NO', '')
REVIEWER     = _c.get('REVIEWER', '')
REVIEW_DATE  = _c.get('REVIEW_DATE') or '%d年%d月%d日' % (datetime.date.today().year,
                                                        datetime.date.today().month,
                                                        datetime.date.today().day)
SECOND_REVIEWER = _c.get('SECOND_REVIEWER', '')
REPLIER      = _c.get('REPLIER', '')
REPLY_DATE   = _c.get('REPLY_DATE', '')
OUTPUT       = _c['OUTPUT']
GROUPS       = _c['GROUPS']
LABEL        = _c.get('LABEL', '上轮意见')
INTRO        = _c.get('INTRO', '')
DOC_TITLE    = _c.get('DOC_TITLE', '审核意见回复落实情况核查表')

PLACEHOLDER = "[项目名称]"
PH_NO, PH_REVIEWER, PH_DATE = "[项目编号]", "[审核人]", "[审核日期]"
TITLE_LINE2_IN_TPL = "三级审核意见及回复记录"     # 模板自带的第2行，本脚本改写为 DOC_TITLE

# ---------------- XML 片段（与 build_opinion.py 保持一致） ----------------
def esc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

RFONTS = ('<w:rFonts w:hint="eastAsia" w:ascii="Times New Roman" w:hAnsi="Times New Roman"'
          ' w:eastAsia="仿宋_GB2312" w:cs="Times New Roman"/>')

def rpr(bold=False, sz=24):
    b = '<w:b/><w:bCs/>' if bold else '<w:b w:val="0"/><w:bCs w:val="0"/>'
    return '<w:rPr>%s%s<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (RFONTS, b, sz, sz)

def run(text, bold=False, sz=24):
    if text == '': return ''
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr(bold, sz), esc(text))

# 序号＝Word 自动编号（同 build_opinion.py，理由见该脚本注释：用户会在 Word 里直接删条）
NUM_ID, ABSTRACT_ID = 20, 20
NUMBERING_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="singleLevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="nothing"/>'
    '<w:lvlText w:val="%%1、"/><w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="0" w:leftChars="0" w:firstLine="0" w:firstLineChars="0"/></w:pPr>'
    '<w:rPr>%s<w:b w:val="0"/><w:bCs w:val="0"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
    '</w:lvl></w:abstractNum>'
    '<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num></w:numbering>'
) % (ABSTRACT_ID, RFONTS, NUM_ID, ABSTRACT_ID)
CT_NUMBERING = ('<Override PartName="/word/numbering.xml" ContentType="application/vnd.'
                'openxmlformats-officedocument.wordprocessingml.numbering+xml"/>')
REL_NUMBERING = ('<Relationship Id="rIdNum%d" Type="http://schemas.openxmlformats.org/'
                 'officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' % NUM_ID)

def para(runs_xml, bdr=False, indent=False, bold_mark=False, num=False):
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
    p.append(rpr(bold_mark)); p.append('</w:pPr>')
    return '<w:p>' + ''.join(p) + runs_xml + '</w:p>'

def cat_header(text):
    return para(run(text, bold=True), bold_mark=True)

def item_block(n, src_no, opinion, reply, check, concl, last=False):
    """一条核查块：N、／上轮意见X：／承做人回复：／落实核查：+正文／核查结论："""
    out = [para('', num=True)]   # 序号段：空文本，「N、」由 Word 自动编号生成（删条自动重排）
    out.append(para(run('%s %s：' % (LABEL, src_no), True) + run(opinion)))
    out.append(para(run('承做人回复：', True) + run(reply)))
    out.append(para(run('落实核查：', True)))
    for seg in [s for s in str(check).split('\n') if s.strip()] or ['']:
        out.append(para(run(seg), indent=True))
    out.append(para(run('核查结论：', True) + run(concl), bdr=True))
    out.append(para('', bdr=True))
    if not last:
        out.append(para(''))
    return ''.join(out)

# ---------------- 结构定位（列数签名） ----------------
def find_tables(doc):
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
        sys.exit('[FAIL] 模板结构异常：%d 列的表格找到 %d 张（应 1 张）。' % (ncols, len(hits)))
    return hits[0]

def strip_runs(xml):
    xml = re.sub(r'<w:r>.*?</w:r>', '', xml, flags=re.S)
    return re.sub(r'<w:r\s[^>]*>.*?</w:r>', '', xml, flags=re.S)

def find_para(doc, needle):
    for m in re.finditer(r'<w:p(?:\s[^>]*)?>.*?</w:p>', doc, re.S):
        if needle in m.group(0): return m
    return None

def fill_info_table(tbl_xml, pairs):
    for label, value in pairs:
        tcs = list(re.finditer(r'<w:tc>.*?</w:tc>', tbl_xml, re.S))
        idx = None
        for i, m in enumerate(tcs):
            if label in re.sub(r'<[^>]+>', '', m.group(0)):
                idx = i; break
        if idx is None or idx + 1 >= len(tcs):
            sys.exit('[FAIL] 模板信息表缺少标签：%s' % label)
        vc = tcs[idx + 1]; raw = vc.group(0)
        m = re.search(r'<w:r>\s*(<w:rPr>.*?</w:rPr>)', raw, re.S)
        vrpr = m.group(1) if m else rpr(False)
        xml = strip_runs(raw)
        if value:
            j = xml.index('</w:p>')
            xml = xml[:j] + '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (vrpr, esc(value)) + xml[j:]
        tbl_xml = tbl_xml[:vc.start()] + xml + tbl_xml[vc.end():]
    return tbl_xml

# ---------------- 组装 ----------------
if not os.path.exists(TEMPLATE):
    sys.exit('[FAIL] 找不到模板：%s' % TEMPLATE)
zin = zipfile.ZipFile(TEMPLATE)
doc = zin.read('word/document.xml').decode('utf-8')

# 1) 标题：第1行替换项目名；第2行由「三级审核意见及回复记录」改写为本表名称
tp = find_para(doc, PLACEHOLDER)
if tp is None:
    sys.exit('[FAIL] 模板标题里找不到占位符 %s。' % PLACEHOLDER)
doc = doc[:tp.start()] + tp.group(0).replace(PLACEHOLDER, esc(PROJECT_NAME)) + doc[tp.end():]
t2 = find_para(doc, TITLE_LINE2_IN_TPL)
if t2 is None:
    sys.exit('[FAIL] 模板标题第2行未找到「%s」。' % TITLE_LINE2_IN_TPL)
doc = (doc[:t2.start()]
       + t2.group(0).replace(TITLE_LINE2_IN_TPL, esc(DOC_TITLE))
       + doc[t2.end():])

# 2) 正文表
spans = find_tables(doc)
if len(spans) != 3:
    sys.exit('[FAIL] 模板应含 3 张表，实际 %d 张。' % len(spans))
body_span = pick_table(doc, spans, 1)
body_xml = doc[body_span[0]:body_span[1]]
m = re.search(r'(<w:tc>)(<w:tcPr>.*?</w:tcPr>)(.*)(</w:tc>)', body_xml, re.S)
if not m:
    sys.exit('[FAIL] 正文表单元格结构异常。')

_groups = [(hd, lst) for hd, lst in GROUPS if lst]
_total = sum(len(lst) for _, lst in _groups)
if _total == 0:
    sys.exit('[FAIL] GROUPS 里没有任何核查条目。')

n = 0; content = []
if INTRO:
    for seg in [s for s in INTRO.split('\n') if s.strip()]:
        content.append(para(run(seg), indent=True))
    content.append(para(''))
for hd, lst in _groups:
    content.append(cat_header(hd))
    for it in lst:
        src_no, opinion, reply, check, concl = (list(it) + ['', '', '', '', ''])[:5]
        n += 1
        content.append(item_block(n, src_no, opinion, reply, check, concl, last=(n == _total)))

new_body = body_xml[:m.start()] + m.group(1) + m.group(2) + ''.join(content) + m.group(4) + body_xml[m.end():]
doc = doc[:body_span[0]] + new_body + doc[body_span[1]:]

# 3) 信息表
spans = find_tables(doc); info_span = pick_table(doc, spans, 2)
info_xml = fill_info_table(doc[info_span[0]:info_span[1]], [
    ('项目编号：', PROJECT_NO), ('二级复核人：', SECOND_REVIEWER), ('审核人：', REVIEWER),
    ('审核日期：', REVIEW_DATE), ('回复人：', REPLIER), ('回复日期：', REPLY_DATE),
])
doc = doc[:info_span[0]] + info_xml + doc[info_span[1]:]

# ---------------- 强制自检 ----------------
errs = []
try: M.parseString(doc.encode('utf-8'))
except Exception as e: errs.append('document.xml 不合法：%s' % e)
plain = re.sub(r'<[^>]+>', '', doc)
plain = plain.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
for ph in (PLACEHOLDER, PH_NO, PH_REVIEWER, PH_DATE):
    if ph in plain: errs.append('占位符 %s 未替换' % ph)
if PROJECT_NAME and PROJECT_NAME[:12] not in plain: errs.append('标题未写入项目名')
if TITLE_LINE2_IN_TPL in plain: errs.append('标题第2行仍是模板原文，未改为本表名称')
if plain.count(DOC_TITLE) != 1: errs.append('表名出现 %d 次（应 1 次）' % plain.count(DOC_TITLE))
_numbered = doc.count('<w:numId w:val="%d"/>' % NUM_ID)
if _numbered != n: errs.append('挂自动编号的序号段 %d 个 ≠ 条数 %d' % (_numbered, n))
if re.search(r'<w:t[^>]*>\d+、</w:t>', doc): errs.append('正文中仍存在字面序号（应改为 Word 自动编号）')
if doc.count('<w:t xml:space="preserve">核查结论：</w:t>') != n:
    errs.append('「核查结论：」段数 ≠ 条数 %d' % n)
_bt = doc[pick_table(doc, find_tables(doc), 1)[0]:pick_table(doc, find_tables(doc), 1)[1]]
if _bt.count('<w:pBdr>') != 2 * n:
    errs.append('正文表内下边框段 %d 个（每条应 2 个，共 %d）' % (_bt.count('<w:pBdr>'), 2 * n))
if errs:
    sys.exit('[FAIL] 自检未通过：\n  - ' + '\n  - '.join(errs))

_ct = zin.read('[Content_Types].xml').decode('utf-8')
if 'numbering+xml' not in _ct: _ct = _ct.replace('</Types>', CT_NUMBERING + '</Types>')
_rels = zin.read('word/_rels/document.xml.rels').decode('utf-8')
if 'relationships/numbering' not in _rels:
    _rels = _rels.replace('</Relationships>', REL_NUMBERING + '</Relationships>')
try: M.parseString(NUMBERING_XML.encode('utf-8')); M.parseString(_ct.encode('utf-8')); M.parseString(_rels.encode('utf-8'))
except Exception as e: sys.exit('[FAIL] 编号部件 XML 不合法：%s' % e)

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
if not _has_numbering: parts.append(('word/numbering.xml', NUMBERING_XML.encode('utf-8')))
target = OUTPUT
try:
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, d in parts: z.writestr(it, d)
except PermissionError:
    target = OUTPUT[:-5] + '（已修订）.docx'
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as z:
        for it, d in parts: z.writestr(it, d)
    print('[锁定→另存]', target)
print('[OK]', target)
print('  标题：对《%s》/ %s' % (PROJECT_NAME, DOC_TITLE))
print('  项目编号 %s | 审核人 %s | 审核日期 %s' % (PROJECT_NO, REVIEWER, REVIEW_DATE))
print('  核查 %d 条，分 %d 个大类' % (n, len(_groups)))
print('  自检通过（XML合法 / 占位符已替 / 表名已改 / %d 条挂自动编号 / 核查结论 %d 段 / 下边框 %d 段）'
      % (n, n, 2 * n))
