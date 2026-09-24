# 估价报告质控审核技能包

[![License: CC BY-NC-SA 4.0](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-lightgrey.svg)](https://creativecommons.org/licenses/by-nc-sa/4.0/deed.zh-hans)

把评估／估价报告的**三级质控审核规则**、**交付物模板**、**构建脚本**打包成一个自包含、可移植的技能包。拷到任何人的电脑、交给任何**能读写文件且能跑 Python** 的 AI 助手，都能按统一标准审核报告，并产出**格式严格一致**的交付物。

> 核心思路：格式一致性**不靠模型模仿样例**，靠「模板克隆 + 脚本生成 + 强制自检」三重约束。规则层（`rules/`）负责"审什么"，脚本层（`scripts/`）负责"长什么样"，两层解耦。

## 覆盖的报告类型

资产评估报告（含企业价值／股权、不动产、年租金等）· 房地产估价报告 · 土地估价报告 · 债权分析报告（金融不良资产）· 矿业权评估报告 · 其他咨询类报告

叠加维度：**司法涉执**（法院委托财产处置参考价）· **执业质量检查模式**（联合检查／自律检查／机构自查口径，与常规审核互斥）

## 快速开始

```bash
# 1. 环境自检（会告诉你缺什么、怎么装）
python scripts/check_env.py

# 2. 抽取待审资料的文本
python scripts/extract.py "报告.docx" --out "审核-项目号-简称/_审核工作文件（勿外发）"
python scripts/extract.py "测算表.xlsx" --out "审核-项目号-简称/_审核工作文件（勿外发）"

# 3. 审完后先出批注版 + 交叉验证，交给人工过目调整批注
python scripts/build_annotated.py 批注配置.json
python scripts/build_crossvalidation.py 交叉验证配置.json

# 4. 人工确认"审核意见明确"后，回读改完的批注 → 生成审核意见（序号现场 1..N 重排）
python scripts/build_opinion_from_annotated.py 意见配置.json          # 加 --preview 只看清单
```

**给 AI 助手的话**：把本文件夹整体给它，让它 ① 先完整读 `SKILL.md`；② 再读 `rules/` 下对应报告类型的要点；③ 严格套 `templates/` + 跑 `scripts/` 产出交付物，**不要自由发挥另起格式**。

## 目录

| 路径 | 内容 |
|---|---|
| `SKILL.md` | **总入口，必先读**。两种工作模式、标准流程、通用规则、交付物规格 |
| `rules/` | 各报告类型的条款级审核要点 + 交付物格式规格 + 公开数据查询站点 |
| `templates/` | Word 模板（**已脱敏的空白模板**） |
| `scripts/` | 资料抽取、交付物生成、环境自检、泄密扫描 |
| `GUIDE-for-colleague.md` | 给同事的上手指南 |
| `cross-platform-checklist.md` | 换电脑／换 AI 平台的一页核对清单 |

## 交付物（**两批交付，中间夹一道人工介入**）

**第一批**
1. **批注版报告** —— Word 真批注，直接描述问题
2. **交叉验证分析报告** —— 数据对比表 + 图表 + 可点击来源链接

**→ 停下来问"审核意见是否明确？"，由人工在 Word 里增删改批注 →**

**第二批**

3. **审核意见** —— 逐条式「位置／原文／审核意见／回复」，按文档大类分组，带回复栏供往复。
   内容**回读自人工确认后的批注版**（`build_opinion_from_annotated.py`）+ 明细表/测算表意见，
   **序号在生成时 1..N 现场重排——删掉任意一条，序号仍然连续。**

`build_opinion.py` 生成时会强制自检 7 类问题（XML 合法性、占位符替换、序号连续、结构完整等），**任一不过直接非零退出、不产出文件**。

## 环境

- Python 3；`openpyxl`（读 .xlsx，缺失自动降级）、`matplotlib` + 中文字体（仅交叉验证出图用）
- docx 处理走标准库 `zipfile`，不依赖 python-docx
- 老式 `.doc` / 扫描版 PDF 读不了 —— 需请报告方转 `.docx` 或提供可读版

## 🔒 保密约定（重要）

**本仓库不含、也不得放入任何真实客户资料、工作底稿或项目成品。** `templates/` 下均为脱敏空白模板。

仓库自带两道防线：

- `.gitignore` —— 默认拒收一切 `.docx/.xls/.pdf` 等文档，仅白名单放行空白模板；屏蔽常见的项目资料目录
- `scripts/check_no_leak.py` —— 泄密扫描，可装成 pre-commit 钩子：

```bash
printf '#!/bin/sh\nexec python scripts/check_no_leak.py --staged\n' > .git/hooks/pre-commit
chmod +x .git/hooks/pre-commit
```

## 许可

[CC BY-NC-SA 4.0](LICENSE) —— 可自由学习、使用、修改与再分发，**须署名、禁止商业使用、演绎作品须以相同协议开放**。

⚠️ 本包是**方法论与工具**，不构成执业结论或法律／评估意见。所引准则、规程、法规的版本与施行日期可能因修订而过时，**使用者须自行核对现行有效版本**，并对据此作出的判断自负其责。
