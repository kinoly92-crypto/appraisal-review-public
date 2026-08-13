# -*- coding: utf-8 -*-
"""给**已含批注**的 docx 追加 Word 真批注（保留原有批注，不覆盖、不重号）。

与 build_annotated.py 的区别：
- build_annotated.py 从零建 comments.xml，若源文档已有批注会被丢弃且 zip 内出现重复条目；
- 本脚本读取源文档已有的 word/comments.xml，把新批注**接在其后**、id 从 max(已有id)+1 起编，
  并保留 commentsExtended.xml 等既有部件。

源文档没有批注时，本脚本等价于 build_annotated.py，可直接通用。

用法（跨平台首选）：python merge_annotated.py 配置.json
JSON 键：SRC / OUTPUT / AUTHOR / INITIALS / DATE(可选) / COMMENTS[[锚点原文, 批注内容], ...]
要点：批注直接描述问题、不加[重大]/[一般]/[提示]前缀；只描述问题，不写依据条款、不写整改建议。
"""
import sys, zipfile, re, json
import xml.dom.minidom as M
try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

# ===================== 【按项目填写】 =====================
SRC    = r"<被批注的报告 docx 绝对路径>"
OUTPUT = r"<输出批注版 docx 绝对路径>"
AUTHOR = "<批注人姓名>"
INITIALS = "<批注人简称>"
import datetime as _dt
DATE = _dt.date.today().isoformat()
COMMENTS = [
    ("<报告中一段原文关键词>", "<批注：只描述问题，不写依据条款和整改建议>"),
]
# =========================================================
if len(sys.argv) > 1:
    _c = json.load(open(sys.argv[1], encoding='utf-8'))
    SRC=_c.get('SRC',SRC); OUTPUT=_c.get('OUTPUT',OUTPUT); AUTHOR=_c.get('AUTHOR',AUTHOR)
    INITIALS=_c.get('INITIALS',INITIALS); COMMENTS=_c.get('COMMENTS',COMMENTS); DATE=_c.get('DATE',DATE)

import html as _html
def esc(s): return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
def _norm(s): return re.sub(r'[\s　 ]', '', _html.unescape(s))

z = zipfile.ZipFile(SRC)
names = z.namelist()
doc = z.read('word/document.xml').decode('utf-8')

# ---- 读取已有批注，确定起始 id ----
old_cx = None
base_id = 0
if 'word/comments.xml' in names:
    old_cx = z.read('word/comments.xml').decode('utf-8')
    ids = [int(x) for x in re.findall(r'<w:comment\b[^>]*\bw:id="(\d+)"', old_cx)]
    base_id = (max(ids) + 1) if ids else 0
    print("源文档已有批注 %d 条，新批注自 id=%d 起编" % (len(ids), base_id))

def ptext(p): return _html.unescape(re.sub(r'<[^>]+>', '', p))
paras = list(re.finditer(r'<w:p\b[^>]*>.*?</w:p>', doc, re.S))
ptxt = [_norm(ptext(m.group(0))) for m in paras]

groups = {}
for k, (kw, _) in enumerate(COMMENTS):
    cid = base_id + k
    nkw = _norm(kw); hit = False
    for i, m in enumerate(paras):
        if nkw and nkw in ptxt[i]:
            groups.setdefault((m.start(), m.end()), []).append(cid); hit = True; break
    if not hit:
        probe = nkw[:8]; cand = [t[:40] for t in ptxt if probe and probe in t][:3]
        print("WARN 未找到锚点:", kw, "| 候选段落:", cand or "无(关键词或跨段/有错字，请改用更短的连续原文)")

repls = []
for (s, e), cids in groups.items():
    ps = doc[s:e]
    st = ''.join('<w:commentRangeStart w:id="%d"/>' % c for c in cids)
    en = ''.join('<w:commentRangeEnd w:id="%d"/><w:r><w:commentReference w:id="%d"/></w:r>' % (c, c) for c in cids)
    inj = ps.replace('</w:pPr>', '</w:pPr>' + st, 1) if '</w:pPr>' in ps else re.sub(r'(<w:p\b[^>]*>)', r'\1' + st, ps, count=1)
    inj = inj[:-len('</w:p>')] + en + '</w:p>'
    repls.append((s, e, inj))
repls.sort(key=lambda x: -x[0])
for s, e, inj in repls: doc = doc[:s] + inj + doc[e:]
print("锚定:", sum(len(v) for v in groups.values()), "/", len(COMMENTS))

new_body = ""
for k, (kw, txt) in enumerate(COMMENTS):
    new_body += ('<w:comment w:id="%d" w:author="%s" w:date="%sT00:00:00Z" w:initials="%s">'
                 '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p></w:comment>'
                 % (base_id + k, esc(AUTHOR), DATE, esc(INITIALS), esc(txt)))

if old_cx:
    # 追加到已有 comments.xml 的 </w:comments> 之前，保留原命名空间声明
    cx = re.sub(r'</w:comments>\s*$', new_body + '</w:comments>', old_cx.rstrip())
else:
    cx = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
          + new_body + '</w:comments>')

ct = z.read('[Content_Types].xml').decode('utf-8')
if 'comments+xml' not in ct:
    ct = ct.replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')
rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
if 'relationships/comments"' not in rels:
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    nid = 'rId%d' % ((max(ids) + 1) if ids else 1)
    rels = rels.replace('</Relationships>', '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>' % nid)

parts = []
for it in z.infolist():
    d = z.read(it.filename)
    if it.filename == 'word/document.xml': d = doc.encode('utf-8')
    elif it.filename == '[Content_Types].xml': d = ct.encode('utf-8')
    elif it.filename == 'word/_rels/document.xml.rels': d = rels.encode('utf-8')
    elif it.filename == 'word/comments.xml': d = cx.encode('utf-8')
    parts.append((it.filename, d))
if 'word/comments.xml' not in names:
    parts.append(('word/comments.xml', cx.encode('utf-8')))

M.parseString(doc); M.parseString(cx)
try:
    with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as zz:
        for fn, d in parts: zz.writestr(fn, d)
    print("[OK]", OUTPUT)
except PermissionError:
    alt = OUTPUT[:-5] + "（已修订）.docx"
    with zipfile.ZipFile(alt, 'w', zipfile.ZIP_DEFLATED) as zz:
        for fn, d in parts: zz.writestr(fn, d)
    print("[锁定→另存]", alt)
