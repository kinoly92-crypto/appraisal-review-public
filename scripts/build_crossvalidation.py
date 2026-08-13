# -*- coding: utf-8 -*-
"""生成「交叉验证分析报告」docx：matplotlib出图 + 克隆模板 + 数据来源可点击超链接。
改下面【按项目填写】区。运行：python build_crossvalidation.py
依赖：模板 ../templates/crossvalidation-template.docx；matplotlib + 中文字体(SimHei)。"""
import sys, zipfile, re, os, tempfile, json
try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import xml.dom.minidom as M

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "templates", "crossvalidation-template.docx")
FONT = r"C:/Windows/Fonts/simhei.ttf"   # 中文字体，按机器调整
if not os.path.exists(FONT):            # 跨平台兜底：自动找一个 CJK 字体
    for _f in [r"C:/Windows/Fonts/msyh.ttc", "/System/Library/Fonts/PingFang.ttc",
               "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"]:
        if os.path.exists(_f): FONT=_f; break

# ===================== 【按项目填写】 =====================
SUBTITLE = "<项目名>（<项目号>）"
OUTPUT   = r"<输出 docx 绝对路径>"
# 数据对比表：表头 + 各行（5列）
TABLE = [
  ["验证维度","报告取值","第三方/印证数据","偏差/勾稽","数据来源"],
  ["<维度>","<取值>","<第三方数据>","<偏差/印证>","<来源>"],
]
# 图1/图2：(标题, [标签...], [数值...])
CHART1 = ("图1  <标题>(元/㎡)", ["A","B","C"], [100,200,300])
CHART2 = ("图2  <标题>", ["甲","乙"], [50,80])
CAP1 = "图1：<题注>"
CAP2 = "图2：<题注>"
CONSISTENCY = ["1. <数据一致性核查条目>", "2. ..."]
# 数据来源：每条 = 文本段列表，元素 ('t',文字) 或 ('l',URL)
SOURCES = [
  [('t','1. <来源名>：'), ('l','https://example.com/'), ('t','（说明）')],
]
CONCLUSION = ["经交叉验证：", "1. ...", "综合判断：建议……"]
# =========================================================
# 跨平台推荐用法：python build_crossvalidation.py 配置.json（用 Write 工具写 UTF-8 JSON，免改脚本/免 heredoc）
if len(sys.argv) > 1:
    _c = json.load(open(sys.argv[1], encoding='utf-8'))
    for _k in ('SUBTITLE','OUTPUT','TABLE','CHART1','CHART2','CAP1','CAP2','CONSISTENCY','SOURCES','CONCLUSION','FONT'):
        if _k in _c: globals()[_k]=_c[_k]

fp = FontProperties(fname=FONT); plt.rcParams['axes.unicode_minus']=False
def mkchart(spec, path):
    title, labels, vals = spec
    fig, ax = plt.subplots(figsize=(7.2,4.0))
    b = ax.bar(labels, vals, width=0.6, color=['#E45756','#4C78A8','#72B7B2','#F0A05A','#BAB0AC'][:len(labels)] or None)
    for r,v in zip(b,vals): ax.text(r.get_x()+r.get_width()/2, v, str(v), ha='center', va='bottom', fontproperties=fp, fontsize=10)
    for l in ax.get_xticklabels(): l.set_fontproperties(fp)
    ax.set_title(title, fontproperties=fp, fontsize=12); ax.grid(axis='y', ls='--', alpha=0.4)
    plt.tight_layout(); plt.savefig(path, dpi=130); plt.close()
c1=os.path.join(tempfile.gettempdir(),"_cv1.png"); c2=os.path.join(tempfile.gettempdir(),"_cv2.png")
mkchart(CHART1,c1); mkchart(CHART2,c2)

def esc(s): return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
F='<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="仿宋_GB2312"/>'
def rpr(b,sz,ex=''): return '<w:rPr>'+F+('<w:b/>' if b else '<w:b w:val="0"/>')+ex+'<w:sz w:val="'+str(sz)+'"/></w:rPr>'
def run(b,sz,t): return '<w:r>'+rpr(b,sz)+'<w:t xml:space="preserve">'+esc(t)+'</w:t></w:r>'
def link(rid,t): return '<w:hyperlink r:id="'+rid+'"><w:r>'+rpr(False,24,'<w:color w:val="0563C1"/><w:u w:val="single"/>')+'<w:t xml:space="preserve">'+esc(t)+'</w:t></w:r></w:hyperlink>'
def pp(ind,jc): return '<w:pPr><w:spacing w:line="360" w:lineRule="auto" w:after="0" w:before="0"/>'+ind+'<w:jc w:val="'+jc+'"/></w:pPr>'
def para(b,sz,t,jc="both",ind='<w:ind w:firstLine="0"/>'): return '<w:p>'+pp(ind,jc)+run(b,sz,t)+'</w:p>'
def title(t): return '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto" w:after="0" w:before="0"/><w:jc w:val="center"/></w:pPr>'+run(True,32,t)+'</w:p>'
def subtitle(t): return '<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto" w:after="240" w:before="0"/><w:jc w:val="center"/></w:pPr>'+run(False,28,t)+'</w:p>'
def heading(t): return para(True,28,t,"both")
def cap(t): return '<w:p><w:pPr><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr>'+run(False,24,t)+'</w:p>'
def bb(t): return para(True,24,t,"both",'<w:ind w:firstLine="420"/>')
def bd(t): return para(False,24,t,"both",'<w:ind w:firstLine="420"/>')
def cell(t,b): return '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1729"/></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'+run(b,24,t)+'</w:p></w:tc>'
def trow(cs,b): return '<w:tr>'+''.join(cell(c,b) for c in cs)+'</w:tr>'
def mixed(segs):
    inner=''; ri=11
    for s in segs:
        if s[0]=='t': inner+=run(False,24,s[1])
    return inner
GRID='<w:tblGrid>'+'<w:gridCol w:w="1729"/>'*5+'</w:tblGrid>'
TBLPR='<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:type="auto" w:w="0"/><w:tblLook w:firstColumn="1" w:firstRow="1" w:lastColumn="0" w:lastRow="0" w:noHBand="0" w:noVBand="1" w:val="04A0"/></w:tblPr>'
table='<w:tbl>'+TBLPR+GRID+trow(TABLE[0],True)+''.join(trow(r,False) for r in TABLE[1:])+'</w:tbl>'
D1='<w:p><w:r><w:drawing><wp:inline xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:extent cx="5040000" cy="2803314"/><wp:docPr id="1" name="P1"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="c1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5040000" cy="2803314"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
D2=D1.replace('rId9','rId10').replace('cy="2803314"','cy="2800946"').replace('id="1" name="P1"','id="2" name="P2"').replace('c1.png','c2.png')
# 数据来源：分配超链接rId（rId11起）
rid_map=[]; rid=11; src_xml=[]
for segs in SOURCES:
    inner=''
    for s in segs:
        if s[0]=='t': inner+=run(False,24,s[1])
        else:
            r='rId%d'%rid; rid_map.append((r,s[1])); rid+=1
            inner+=link(r,s[1])
    src_xml.append('<w:p>'+pp('<w:ind w:firstLine="420"/>','both')+inner+'</w:p>')

bx=title("交叉验证分析报告")+subtitle(SUBTITLE)
bx+=heading("一、数据对比表")+table
bx+=heading("二、数据对比可视化")+cap(CAP1)+D1+cap(CAP2)+D2
bx+=heading("三、分析说明")+bb("【数据一致性核查】")+"".join(bd(t) for t in CONSISTENCY)+bb("【数据来源说明】")+"".join(src_xml)
bx+=heading("四、验证结论")+bb("【验证结论】")+"".join(bd(t) for t in CONCLUSION)

doc=zipfile.ZipFile(TEMPLATE).read('word/document.xml').decode('utf-8')
pre=doc[:doc.index('<w:body>')+len('<w:body>')]; sect=re.search(r'<w:sectPr\b.*?</w:sectPr>',doc,re.S).group(0)
new_doc=pre+bx+sect+'</w:body></w:document>'
rels=zipfile.ZipFile(TEMPLATE).read('word/_rels/document.xml.rels').decode('utf-8')
rels=re.sub(r'<Relationship Id="rId1[1-9]"[^>]*/>','',rels)
rels=rels.replace('</Relationships>',''.join('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>'%(r,esc(u)) for r,u in rid_map)+'</Relationships>')
M.parseString(new_doc); M.parseString(rels)
parts=[]
for it in zipfile.ZipFile(TEMPLATE).infolist():
    d=zipfile.ZipFile(TEMPLATE).read(it.filename)
    if it.filename=='word/document.xml': d=new_doc.encode('utf-8')
    elif it.filename=='word/_rels/document.xml.rels': d=rels.encode('utf-8')
    elif it.filename=='word/media/image1.png': d=open(c1,'rb').read()
    elif it.filename=='word/media/image2.png': d=open(c2,'rb').read()
    parts.append((it,d))
try:
    with zipfile.ZipFile(OUTPUT,'w',zipfile.ZIP_DEFLATED) as z:
        for it,d in parts: z.writestr(it,d)
    print("[OK]", OUTPUT, "| 超链接", len(rid_map))
except PermissionError:
    alt=OUTPUT[:-5]+"（已修订）.docx"
    with zipfile.ZipFile(alt,'w',zipfile.ZIP_DEFLATED) as z:
        for it,d in parts: z.writestr(it,d)
    print("[锁定→另存]", alt)
