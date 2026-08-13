# -*- coding: utf-8 -*-
"""环境自检：一次性确认依赖，避免每次会话重复 pip install。
用法：python check_env.py
原则：缺啥才给唯一一行安装命令；都齐了就直接出三件套，本会话内别再装。"""
import sys, os, importlib
try: sys.stdout.reconfigure(encoding='utf-8')
except Exception: pass

print("Python:", sys.version.split()[0])
OPT = {
    'openpyxl':   '读 .xlsx 保真（缺失会自动用纯标准库降级读取，仍可用，非必装）',
    'matplotlib': '交叉验证出图必需（仅“涉及估值”的项目用到；纯批注+审核意见不需要）',
}
miss = []
for m, why in OPT.items():
    try:
        importlib.import_module(m); print("[OK] %-11s — %s" % (m, why))
    except Exception:
        print("[缺] %-11s — %s" % (m, why)); miss.append(m)

fonts = [r"C:/Windows/Fonts/simhei.ttf", r"C:/Windows/Fonts/msyh.ttc",
         "/System/Library/Fonts/PingFang.ttc",
         "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"]
hit = next((p for p in fonts if os.path.exists(p)), None)
print("中文字体:", hit or "未在常见路径找到（交叉验证出图需要，可改 build_crossvalidation.py 顶部 FONT）")

if miss:
    print("\n一次性安装（只需一次，本会话内勿重复）：\n    pip install " + " ".join(miss))
else:
    print("\n依赖齐全，直接出三件套，无需任何 pip install。")
