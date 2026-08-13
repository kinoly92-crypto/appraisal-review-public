# -*- coding: utf-8 -*-
"""泄密自检：扫描仓库（或 git 暂存区）里有没有真实项目信息。

用法：
    python scripts/check_no_leak.py            # 扫全部被 git 跟踪的文件
    python scripts/check_no_leak.py --staged   # 只扫 git 暂存区（供 pre-commit 钩子用）

装成提交前钩子（一次性）：
    printf '#!/bin/sh\\nexec python scripts/check_no_leak.py --staged\\n' > .git/hooks/pre-commit
    chmod +x .git/hooks/pre-commit

命中即以非零码退出，阻止提交。确属误报时用 git commit --no-verify 跳过（慎用）。
"""
import sys, os, re, subprocess, zipfile

try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

# —— 强信号：出现即判定泄密 ——
# 主体类模式一律要求前面挂着 ≥2 个汉字的"名字"，否则"有限责任公司""分行"这类
# 法条/方法术语会误报（例：《公司法》条文、"分层抽样(分行/…)"）。
CJK = r'[一-龥]'
HARD = [
    CJK + r'{2,}(?:有限公司|有限责任公司|股份有限公司|合作社)',
    CJK + r'{2,}(?:支行|分行)',
    r'不动产权第\s*\d+', r'不动产权证书', r'国用[（(]\d{4}[)）]',
    r'评报字[（(]\d{4}[)）]', r'房估报字', r'矿评报字', r'[一-龥]{2,}评报字',
    r'统一社会信用代码', r'身份证号',
    r'法定代表人[：:]\s*' + CJK + r'{2,}',
    r'权利人[：:]\s*' + CJK + r'{3,}',
]
# —— 弱信号：需与其他信号同现才报 ——
SOFT = [r'万元', r'元/㎡', r'元/平方米', r'评估价值', r'估价结果', r'基准日', r'价值时点']

# 白名单只放行"扫描器自身"和"明确标注为占位/脱敏"的行。
# ⚠️ 切勿按"准则/办法/指南"等词放行整行——审核意见正文句句带准则名，
#    那样会把真实客户信息一并放过（本脚本初版就栽在这里，阳性测试 0 命中）。
ALLOW_FILES = {'scripts/check_no_leak.py', '.gitignore'}
ALLOW_LINE = re.compile(r'××|＜占位|<占位|\[项目名称\]|\[项目编号\]|\[审核人\]|\[审核日期\]'
                        r'|此处为审核意见正文|noleak')   # 行内写 noleak 可显式豁免（需人工确认过）

DOC_EXT = ('.docx', '.doc', '.xlsx', '.xls', '.pdf', '.wps', '.et')


def tracked_files(staged):
    cmd = ['git', 'diff', '--cached', '--name-only', '--diff-filter=ACM'] if staged \
        else ['git', 'ls-files']
    out = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    return [f.strip() for f in out.stdout.splitlines() if f.strip()]


def docx_text(path):
    try:
        z = zipfile.ZipFile(path)
        s = ''
        for n in z.namelist():
            if n.endswith('.xml'):
                s += z.read(n).decode('utf-8', 'ignore')
        return re.sub(r'<[^>]+>', '', s)
    except Exception:
        return ''


def scan(path):
    """返回该文件的命中列表 [(行号, 命中词, 行摘要)]。"""
    if path in ALLOW_FILES or not os.path.exists(path):
        return []
    low = path.lower()
    if low.endswith(('.docx', '.xlsx')):
        lines = [(0, docx_text(path))]
    elif low.endswith(DOC_EXT):
        return [(0, '二进制文档', '禁止提交 %s 类文件（.gitignore 应已拦截）' % os.path.splitext(path)[1])]
    else:
        try:
            with open(path, encoding='utf-8', errors='ignore') as fh:
                lines = list(enumerate(fh, 1))
        except Exception:
            return []
    hits = []
    for ln, text in lines:
        if ALLOW_LINE.search(text or ''):
            continue
        hard = [p for p in HARD if re.search(p, text or '')]
        soft = [p for p in SOFT if re.search(p, text or '')]
        if hard or len(soft) >= 3:
            kw = (hard + soft)[:3]
            snippet = re.sub(r'\s+', ' ', (text or ''))[:90]
            hits.append((ln, '/'.join(kw), snippet))
    return hits


def main():
    staged = '--staged' in sys.argv
    files = tracked_files(staged)
    if not files:
        print('[check_no_leak] 没有待检文件'); return 0
    total = 0
    for f in files:
        for ln, kw, snip in scan(f):
            total += 1
            loc = ('%s:%d' % (f, ln)) if ln else f
            print('  [泄密疑似] %-46s %-14s %s' % (loc, kw, snip))
    print()
    if total:
        print('[check_no_leak] ✗ 命中 %d 处疑似真实项目信息。' % total)
        print('  本仓库对外公开，请删除/脱敏后再提交；确属误报可用 git commit --no-verify 跳过。')
        return 1
    print('[check_no_leak] ✓ 扫描 %d 个文件，未发现真实项目信息。' % len(files))
    return 0


if __name__ == '__main__':
    sys.exit(main())
