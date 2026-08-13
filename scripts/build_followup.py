# -*- coding: utf-8 -*-
"""生成「一审审核意见回复落实情况核查表」docx：克隆审核意见模板样式，逐条四段块。
用法：python build_followup.py 配置.json"""
import sys, zipfile, re, os, json
import xml.dom.minidom as M
try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "templates", "opinion-template.docx")

_c = json.load(open(sys.argv[1], encoding='utf-8'))
TITLE=_c['TITLE']; PROJECT_NO=_c['PROJECT_NO']; REVIEWER=_c['REVIEWER']
REVIEW_DATE=_c['REVIEW_DATE']; OUTPUT=_c['OUTPUT']; INTRO=_c['INTRO']; GROUPS=_c['GROUPS']
LABEL=_c.get('LABEL', "一审意见")   # 上一轮意见的称谓：一审意见/二审意见/专家评审意见…

def esc(s): return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
F='<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="仿宋_GB2312"/>'
def rpr(b,sz): return '<w:rPr>'+F+('<w:b/>' if b else '<w:b w:val="0"/>')+'<w:sz w:val="'+str(sz)+'"/></w:rPr>'
def run(b,sz,t): return '<w:r>'+rpr(b,sz)+'<w:t xml:space="preserve">'+esc(t)+'</w:t></w:r>'
PP='<w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="both"/></w:pPr>'
PPI='<w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="420"/><w:jc w:val="both"/></w:pPr>'
PPS='<w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr>'
EMPTY='<w:p>'+PP+'</w:p>'
NUMP='<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="both"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="50"/></w:numPr><w:rPr>'+F+'<w:b/><w:sz w:val="24"/></w:rPr></w:pPr></w:p>'
def header(t): return '<w:p><w:pPr><w:spacing w:before="120" w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="both"/></w:pPr>'+run(True,24,t)+'</w:p>'
def plc(l,c): return '<w:p>'+PP+run(True,24,l)+run(False,24,c)+'</w:p>'
def pop(t): return '<w:p>'+PPI+run(False,24,t)+'</w:p>'
def pintro(t): return '<w:p>'+PPI+run(False,24,t)+'</w:p>'
def psep(): return '<w:p>'+PPS+run(False,24,"——————————————————————————")+'</w:p>'
def block(no,op,rep,chk,concl):
    return ('<w:p>'+PP+run(True,24,LABEL+" "+no+"：")+run(False,24,op)+'</w:p>'
        + plc("承做人回复：", rep)
        + '<w:p>'+PP+run(True,24,"落实核查：")+'</w:p>'
        + pop(chk)
        + plc("核查结论：", concl)
        + psep())

def _settext(rxml, v):
    a=rxml.index('>')+1; b=rxml.rindex('</w:t>'); return rxml[:a]+esc(v)+rxml[b:]
def set_after(doc, label, value):
    tcs=list(re.finditer(r'<w:tc\b.*?</w:tc>', doc, re.S))
    for i,m in enumerate(tcs):
        if label in re.sub(r'<[^>]+>','', m.group(0)) and i+1<len(tcs):
            vc=tcs[i+1]; xml=vc.group(0); runs=list(re.finditer(r'<w:t(?:\s[^>]*)?>.*?</w:t>', xml, re.S))
            if runs:
                for j in range(len(runs)-1,-1,-1):
                    r=runs[j]; xml=xml[:r.start()]+_settext(r.group(0), value if j==0 else '')+xml[r.end():]
            else:
                xml=xml.replace('</w:p>', '<w:r><w:rPr>'+F+'<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">'+esc(value)+'</w:t></w:r></w:p>', 1)
            return doc[:vc.start()]+xml+doc[vc.end():]
    print("WARN 模板信息表未找到标签:", label); return doc

zin=zipfile.ZipFile(TEMPLATE); doc=zin.read('word/document.xml').decode('utf-8')
old_title=re.search(r'对《[^》]*》审核意见', doc)
if old_title: doc=doc.replace(old_title.group(0), esc(TITLE), 1)
doc=doc.replace("2026REA0278", esc(PROJECT_NO))
doc=set_after(doc, "审核人：", REVIEWER)
doc=set_after(doc, "审核日期：", REVIEW_DATE)
i_tbl=doc.index("</w:tbl>")+len("</w:tbl>"); i_sect=doc.index("<w:sectPr")
body=EMPTY + "".join(pintro(t) for t in INTRO) + EMPTY
for hd,lst in GROUPS: body+=header(hd)+"".join(block(*it) for it in lst)
body+=EMPTY
doc=doc[:i_tbl]+body+doc[i_sect:]
nx=zin.read('word/numbering.xml').decode('utf-8')
ABS='<w:abstractNum w:abstractNumId="100"><w:nsid w:val="3A3A3A3A"/><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="nothing"/><w:lvlText w:val="%1："/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr><w:rPr>'+F+'<w:b/><w:sz w:val="24"/></w:rPr></w:lvl></w:abstractNum>'
NUM='<w:num w:numId="50"><w:abstractNumId w:val="100"/></w:num>'
m=re.search(r'<w:num\b', nx); nx=nx[:m.start()]+ABS+nx[m.start():]; nx=nx.replace('</w:numbering>', NUM+'</w:numbering>')
M.parseString(doc); M.parseString(nx)
parts=[]
for it in zin.infolist():
    d=zin.read(it.filename)
    if it.filename=='word/document.xml': d=doc.encode('utf-8')
    elif it.filename=='word/numbering.xml': d=nx.encode('utf-8')
    parts.append((it,d))
try:
    with zipfile.ZipFile(OUTPUT,'w',zipfile.ZIP_DEFLATED) as z:
        for it,d in parts: z.writestr(it,d)
    print("[OK]", OUTPUT, "| numPr", doc.count('<w:numId w:val="50"/>'))
except PermissionError:
    alt=OUTPUT[:-5]+"（已修订）.docx"
    with zipfile.ZipFile(alt,'w',zipfile.ZIP_DEFLATED) as z:
        for it,d in parts: z.writestr(it,d)
    print("[锁定→另存]", alt)
