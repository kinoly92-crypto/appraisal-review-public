# -*- coding: utf-8 -*-
"""资料读取助手：docx/xlsx 文本提取、PDF 文字层探测、老.doc UTF-16启发式。
用法示例：
    python extract.py "某报告.docx" --out "<审核目录>/_审核工作文件（勿外发）"   # -> 该目录下同名.txt
    python extract.py "测算表.xlsx" --out "<工作目录>"   # 非隐藏sheet(隐藏/极隐藏自动跳过)；缺省 --out = 当前目录
    （送审文件夹只读：输出绝不写回送审文件夹）
    python extract.py --probe "某.pdf"      # 探测PDF是否有文字层
说明：老式OLE2二进制.doc多半读不出中文，请让报告方另存为.docx。
"""
import sys, os, re, zipfile, html
try: sys.stdout.reconfigure(encoding='utf-8')   # Win 控制台中文不崩
except Exception: pass

def docx_text(f):
    x = zipfile.ZipFile(f).read('word/document.xml').decode('utf-8','ignore')
    x = re.sub(r'</w:p>', '\n', x); x = html.unescape(re.sub(r'<[^>]+>', '', x))
    x = re.sub(r'[ \t\xa0]+', ' ', x)
    return re.sub(r'\n{2,}', '\n', x).strip()

def xlsx_text(f):
    """有 openpyxl 用它(保真)；没有则纯标准库降级读取——无需任何 pip 安装。"""
    try:
        import openpyxl
    except Exception:
        return _xlsx_stdlib(f)
    try:
        wb = openpyxl.load_workbook(f, data_only=True)
    except Exception:
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        if getattr(ws, 'sheet_state', 'visible') != 'visible':
            continue   # 跳过隐藏/极隐藏 sheet，只审非隐藏表格
        out.append("\n==== SHEET: %s ====" % ws.title)
        for row in ws.iter_rows():
            cells = []
            for c in row:
                v = c.value
                if v is not None and str(v).strip() != '': cells.append(str(v).strip())
            if cells: out.append(" | ".join(cells))
    return "\n".join(out)

def _xlsx_stdlib(f):
    """纯标准库读 .xlsx(zip+XML)：免装 openpyxl。公式/日期取缓存值，日期可能显示为序列号。"""
    import xml.etree.ElementTree as ET
    Z = zipfile.ZipFile(f); names = Z.namelist()
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    rns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    shared = []
    if 'xl/sharedStrings.xml' in names:
        for si in ET.fromstring(Z.read('xl/sharedStrings.xml')).findall(ns+'si'):
            shared.append(''.join(t.text or '' for t in si.iter(ns+'t')))
    rid2t = {r.get('Id'): r.get('Target') for r in ET.fromstring(Z.read('xl/_rels/workbook.xml.rels'))}
    sheets = []
    for sh in ET.fromstring(Z.read('xl/workbook.xml')).find(ns+'sheets'):
        if (sh.get('state') or 'visible') != 'visible':
            continue   # 跳过隐藏/极隐藏 sheet(state=hidden/veryHidden)，只审非隐藏表格
        tgt = rid2t.get(sh.get(rns+'id'), '')
        if tgt.startswith('/'): tgt = tgt.lstrip('/')        # 绝对：从包根
        elif tgt and not tgt.startswith('xl/'): tgt = 'xl/' + tgt   # 相对：基于 xl/
        sheets.append((sh.get('name'), tgt))
    out = []
    for name, path in sheets:
        if path not in names: continue
        out.append("\n==== SHEET: %s ====" % name)
        data = ET.fromstring(Z.read(path)).find(ns+'sheetData')
        if data is None: continue
        for row in data.findall(ns+'row'):
            cells = []
            for c in row.findall(ns+'c'):
                t = c.get('t'); v = c.find(ns+'v')
                if t == 'inlineStr':
                    iss = c.find(ns+'is'); val = ''.join(x.text or '' for x in iss.iter(ns+'t')) if iss is not None else ''
                elif v is None or v.text is None: val = ''
                elif t == 's': val = shared[int(v.text)] if v.text.isdigit() and int(v.text) < len(shared) else ''
                else: val = v.text
                val = (val or '').strip()
                if val: cells.append(val)
            if cells: out.append(" | ".join(cells))
    return "\n".join(out)

def doc_utf16(f):
    """老.doc UTF-16启发式：多为噪声，仅供应急；建议转docx。"""
    d = open(f, 'rb').read().decode('utf-16-le', 'ignore')
    runs = re.findall(r'[一-鿿　-〿＀-￯0-9A-Za-z（）()、。，：；％%\.\-—／/㎡²\s]{2,}', d)
    t = '\n'.join(r.strip() for r in runs if re.search(r'[一-鿿]', r) and len(r.strip()) >= 2)
    return re.sub(r'\n{2,}', '\n', t)

def probe_pdf(f):
    import subprocess
    import tempfile
    tmp = os.path.join(tempfile.gettempdir(), "__probe_%d.txt" % os.getpid())   # 不写进送审文件夹
    os.system('pdftotext -f 1 -l 8 -enc UTF-8 "%s" "%s"' % (f, tmp))
    n = os.path.getsize(tmp) if os.path.exists(tmp) else 0
    if os.path.exists(tmp): os.remove(tmp)
    return n

if __name__ == '__main__':
    args = sys.argv[1:]
    # 输出目录：--out 指定（应为 审核-<项目号>-<简称>\_审核工作文件（勿外发）），缺省=当前工作目录。
    # 送审文件夹只读，任何中间文件都不得写回送审文件夹。
    outdir = None
    if '--out' in args:
        i = args.index('--out'); outdir = args[i + 1]; del args[i:i + 2]
    if args and args[0] == '--probe':
        n = probe_pdf(args[1]); print("PDF前8页文字层 %d 字节 -> %s" % (n, "有文字层可读" if n > 100 else "扫描件/无文字层，需OCR或转换"))
    else:
        f = args[0]; low = f.lower()
        if low.endswith('.docx'): t = docx_text(f)
        elif low.endswith('.xlsx'): t = xlsx_text(f)
        elif low.endswith('.doc'): t = doc_utf16(f)
        else: print("不支持的格式"); sys.exit(1)
        outdir = outdir or os.getcwd()
        os.makedirs(outdir, exist_ok=True)
        out = os.path.join(outdir, os.path.splitext(os.path.basename(f))[0] + ".txt")
        if os.path.abspath(os.path.dirname(out)) == os.path.abspath(os.path.dirname(f)):
            print("[WARN] 输出落在送审文件夹内，违反“送审只读”规则，请改用 --out 指向 _审核工作文件（勿外发）")
        open(out, 'w', encoding='utf-8').write(t)
        print("CJK字:", sum(len(s) for s in re.findall(r'[一-鿿]+', t)), "-> ", out)
