# -*- coding: utf-8 -*-
"""给报告 docx 注入 Word 真批注（从零建 comments.xml + 锚点 + rels + 内容类型）。
改下面【按项目填写】区。运行：python build_annotated.py
要点：批注直接描述问题、不加[重大]/[一般]/[提示]前缀；作者=本次批注人姓名。
锚点用关键词匹配所在段落；同段可挂多条。"""
import sys, zipfile, re, os, json
import xml.dom.minidom as M
try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

# ===================== 【按项目填写】 =====================
SRC    = r"<被批注的报告 docx 绝对路径>"     # 资产评估须分别对 报告 和 说明 各跑一次
OUTPUT = r"<输出批注版 docx 绝对路径>"
AUTHOR = "<批注人姓名>"                        # 本次问询所得，勿写死
INITIALS = "<批注人简称>"
import datetime as _dt
DATE = _dt.date.today().isoformat()            # 批注日期，默认今天；可在 JSON 用 DATE 覆盖(YYYY-MM-DD)
# 批注列表：(锚点关键词=报告原文中能唯一定位的一段连续文字, 批注内容)
COMMENTS = [
    ("<报告中一段原文关键词>", "<批注：只描述问题，不写依据条款和整改建议>"),
]
# =========================================================
# 跨平台推荐用法：python build_annotated.py 配置.json（用 Write 工具写 UTF-8 JSON，免改脚本/免 heredoc）
if len(sys.argv) > 1:
    _c = json.load(open(sys.argv[1], encoding='utf-8'))
    SRC=_c.get('SRC',SRC); OUTPUT=_c.get('OUTPUT',OUTPUT); AUTHOR=_c.get('AUTHOR',AUTHOR)
    INITIALS=_c.get('INITIALS',INITIALS); COMMENTS=_c.get('COMMENTS',COMMENTS); DATE=_c.get('DATE',DATE)

import html as _html
def esc(s): return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
def _norm(s): return re.sub(r'[\s 　]', '', _html.unescape(s))   # 反转义+去所有空白，稳健匹配
z=zipfile.ZipFile(SRC); doc=z.read('word/document.xml').decode('utf-8')
def ptext(p): return _html.unescape(re.sub(r'<[^>]+>','',p))
paras=list(re.finditer(r'<w:p\b[^>]*>.*?</w:p>', doc, re.S))
ptxt=[_norm(ptext(m.group(0))) for m in paras]
groups={}
for cid,(kw,_) in enumerate(COMMENTS):
    nkw=_norm(kw); hit=False
    for i,m in enumerate(paras):
        if nkw and nkw in ptxt[i]:
            groups.setdefault((m.start(),m.end()),[]).append(cid); hit=True; break
    if not hit:
        probe=nkw[:8]; cand=[t[:40] for t in ptxt if probe and probe in t][:3]
        print("WARN 未找到锚点:", kw, "| 候选段落:", cand or "无(关键词或跨段/有错字，请改用更短的连续原文)")
repls=[]
for (s,e),cids in groups.items():
    ps=doc[s:e]
    st=''.join('<w:commentRangeStart w:id="%d"/>'%c for c in cids)
    en=''.join('<w:commentRangeEnd w:id="%d"/><w:r><w:commentReference w:id="%d"/></w:r>'%(c,c) for c in cids)
    inj=ps.replace('</w:pPr>','</w:pPr>'+st,1) if '</w:pPr>' in ps else re.sub(r'(<w:p\b[^>]*>)',r'\1'+st,ps,count=1)
    inj=inj[:-len('</w:p>')]+en+'</w:p>'
    repls.append((s,e,inj))
repls.sort(key=lambda x:-x[0])
for s,e,inj in repls: doc=doc[:s]+inj+doc[e:]
print("锚定:", sum(len(v) for v in groups.values()), "/", len(COMMENTS))
cbody=""
for cid,(kw,txt) in enumerate(COMMENTS):
    cbody+='<w:comment w:id="%d" w:author="%s" w:date="%sT00:00:00Z" w:initials="%s"><w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p></w:comment>'%(cid,esc(AUTHOR),DATE,esc(INITIALS),esc(txt))
cx='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+cbody+'</w:comments>'
ct=z.read('[Content_Types].xml').decode('utf-8')
if 'comments+xml' not in ct: ct=ct.replace('</Types>','<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')
rels=z.read('word/_rels/document.xml.rels').decode('utf-8')
ids=[int(x) for x in re.findall(r'Id="rId(\d+)"',rels)]; nid='rId%d'%((max(ids)+1) if ids else 1)
rels=rels.replace('</Relationships>','<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'%nid)
parts=[]
for it in z.infolist():
    d=z.read(it.filename)
    if it.filename=='word/document.xml': d=doc.encode('utf-8')
    elif it.filename=='[Content_Types].xml': d=ct.encode('utf-8')
    elif it.filename=='word/_rels/document.xml.rels': d=rels.encode('utf-8')
    parts.append((it.filename,d))
parts.append(('word/comments.xml',cx.encode('utf-8')))
M.parseString(doc); M.parseString(cx)
try:
    with zipfile.ZipFile(OUTPUT,'w',zipfile.ZIP_DEFLATED) as zz:
        for fn,d in parts: zz.writestr(fn,d)
    print("[OK]", OUTPUT)
except PermissionError:
    alt=OUTPUT[:-5]+"（已修订）.docx"
    with zipfile.ZipFile(alt,'w',zipfile.ZIP_DEFLATED) as zz:
        for fn,d in parts: zz.writestr(fn,d)
    print("[锁定→另存]", alt)
