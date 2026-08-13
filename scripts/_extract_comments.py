# -*- coding: utf-8 -*-
"""从批注版 docx 抽取全部批注：id/author/text + 锚定段落原文 + 所属章节标题，按文档顺序输出 JSON。
用法：python _extract_comments.py <批注版.docx> <文档大类标签>
"""
import sys, zipfile, re, html, json
try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
import xml.etree.ElementTree as ET

def q(tag): return '{%s}%s' % (W, tag)

def para_text(p):
    """段落纯文本"""
    out = []
    for t in p.iter(q('t')):
        out.append(t.text or '')
    return ''.join(out).strip()

def is_heading(p):
    """返回 (是否标题, 文本) —— 用 pStyle 含 Heading 或 大纲级别判断"""
    ppr = p.find(q('pPr'))
    if ppr is None: return False
    st = ppr.find(q('pStyle'))
    if st is not None:
        v = st.get(q('val')) or ''
        if 'Heading' in v or v.startswith('1') or '标题' in v:
            return True
    return False

def main(path, label):
    Z = zipfile.ZipFile(path)
    names = Z.namelist()
    # 1. 读批注正文
    comments = {}
    if 'word/comments.xml' in names:
        root = ET.fromstring(Z.read('word/comments.xml'))
        for c in root.findall(q('comment')):
            cid = c.get(q('id'))
            txt = ''.join(t.text or '' for t in c.iter(q('t'))).strip()
            comments[cid] = {'author': c.get(q('author')), 'text': txt}
    # 2. 遍历 document.xml，按顺序找 commentRangeStart，记录所在段落文本 + 最近标题
    doc = ET.fromstring(Z.read('word/document.xml'))
    body = doc.find(q('body'))
    cur_heading = ''
    results = []
    # 需要按文档顺序遍历所有段落
    def walk(el):
        nonlocal cur_heading
        for child in list(el):
            tag = child.tag
            if tag == q('p'):
                txt = para_text(child)
                if is_heading(child) and txt:
                    cur_heading = txt
                # 该段内的 commentRangeStart
                for crs in child.iter(q('commentRangeStart')):
                    cid = crs.get(q('id'))
                    results.append({'id': cid, 'heading': cur_heading, 'anchor': txt})
            else:
                walk(child)
    walk(body)
    # 3. 合并
    merged = []
    for r in results:
        c = comments.get(r['id'], {})
        merged.append({
            'label': label,
            'id': r['id'],
            'heading': r['heading'],
            'anchor': r['anchor'],
            'author': c.get('author', ''),
            'comment': c.get('text', ''),
        })
    return merged

if __name__ == '__main__':
    path = sys.argv[1]; label = sys.argv[2]; outpath = sys.argv[3]
    out = main(path, label)
    with open(outpath, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print('wrote %d comments -> %s' % (len(out), outpath))
