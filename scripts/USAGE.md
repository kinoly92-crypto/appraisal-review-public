# 脚本使用说明

所有脚本用 Python 运行（需能读写文件）。建议在本技能包目录或项目目录下执行。

## 依赖（先跑一次 `python check_env.py`，装一次别重复装）
- Python 3。**依赖在同一会话内持久，缺了才装、且只装一次。**
- `openpyxl`（读 .xlsx 测算表）——**可选**：缺失时 `extract.py` 自动用纯标准库降级读取，仍可用（日期可能显示为序列号）。
- `matplotlib` + 中文字体——**仅交叉验证出图用到**（脚本默认 `C:/Windows/Fonts/simhei.ttf`，找不到自动兜底找其他 CJK 字体；也可在 JSON 配置里给 `FONT`）。纯批注+审核意见两件套不需要。
- docx 处理用内置 `zipfile`，无需 python-docx。
- 一键安装（可选增强）：`pip install -r requirements.txt`。

## 流程
1. **读资料**：`python extract.py "报告.docx"`、`python extract.py "测算表.xlsx"`、`python extract.py --probe "扫描件.pdf"`。
   - 老 `.doc` 读不出中文 → 请报告方另存 `.docx`。扫描 PDF 无文字层 → 需 OCR 或请对方提供可读版/口述关键信息。
2. **审核**：按 `../rules/` 对应报告类型逐项核对，**先核测算表再定性**。
3. **出三件套——分两批，中间等人工**（先问到审核人/批注人姓名）：
   **第一批** `build_annotated.py`（批注版，报告/说明各一份）+ `build_crossvalidation.py`（交叉验证）
   → **停下来问"审核意见是否明确？"，等用户在 Word 里增删改批注并回复**
   → **第二批** `build_opinion_from_annotated.py`（回读改完后的批注 + 合并明细表/测算表意见 → 审核意见 docx，序号现场 1..N 重排）。
   **填数据两种方式，跨平台首选 JSON：**
   - **方式A（跨平台首选，免乱码）**：用 **Write 工具**写一个 UTF-8 的 `配置.json`（键名 = 脚本里大写变量名），再 `python build_xxx.py 配置.json`。**不要用 heredoc / `python -c` 把中文塞进 shell**（Windows Git Bash 会乱码、脚本失败）。
   - 方式B：直接改脚本顶部"【按项目填写】"区 → `python build_xxx.py`。
   - `build_opinion.py`（**v2 模板·三级审核意见及回复记录·索引号 G-18-3**）：键 `PROJECT_NAME`/`PROJECT_NO`/`REVIEWER`/`REVIEW_DATE`/`SECOND_REVIEWER`/`REPLIER`/`REPLY_DATE`/`OUTPUT`/`GROUPS`。
     - `GROUPS = [[类别提示, [[位置, 原文, 审核意见], …]], …]`；位置/原文可留空串（该行不输出）；审核意见含 `\n` 自动拆多个缩进段；**空大类自动跳过**。
     - 信息表（项目编号/二级复核人/审核人/审核日期/回复人/回复日期）由脚本自动写入，**无需手改**；文末签字表保持空白待手签。
     - 序号＝**Word 自动编号**（脚本自动注入 numbering.xml），跨大类连续。**在 Word 里直接删条会自动重排、不断号**；改配置重跑脚本亦然。勿手打序号。
     - 收尾强制自检：XML 合法 / 占位符已替 / 自动编号段数＝条数且无字面序号 / 信息表已写入 / 三表完整。**报 `[FAIL]` 即不产出文件**，请按提示查模板是否被改动，勿绕过自检手工拼装。
     - 旧配置的 `TITLE="对《X》审核意见"` 仍兼容（自动提取项目名）。
   - `build_opinion_from_annotated.py`（**新流程主力：人工确认批注后出审核意见**）：键 `PROJECT_*`/`REVIEWER`/`REVIEW_DATE`/`OUTPUT`（同 build_opinion）＋ `SOURCES`/`EXTRA_GROUPS`/`DROP`/`MAX_QUOTE`/`REPLY_MODE`/`INCLUDE_DONE`/`DUMP_JSON`。
     - `SOURCES=[{"LABEL":"资产评估报告：","DOCX":"…批注版.docx","AUTHORS":["批注人姓名"],"POSITION_FALLBACK":"资产评估报告正文"}]`——数组顺序＝大类顺序；`AUTHORS` 白名单**建议必填**（滤掉承做人留在源文档里的旧批注）。
     - 「位置」自动取章节标题（第X章／一、／（一）／1.1／Heading 样式，两级拼接）；「原文」取批注锚定范围原文（默认截 60 字，`MAX_QUOTE:0` 不截）；「审核意见」＝批注正文原样。
     - 默认**跳过** Word 里标"已解决"的批注（`INCLUDE_DONE`）与回复型子批注（`REPLY_MODE`）；过滤/跳过条数都会打印，便于核对。
     - `python build_opinion_from_annotated.py 配置.json --preview` 先看逐条清单不产文件；`--dump 批注版.docx` 只看某份文档现存批注。
     - `DROP:[3,"关键词"]` 再剔任意条 → **序号自动重排，永远连续**；不要手改编号。
     - 会把合并后的 build_opinion 配置落盘（`DUMP_JSON`，默认 `_审核意见_合并配置.json`），人工想再改文字就改它、直接重跑 `build_opinion.py`。
   - `build_annotated.py`：键 SRC/OUTPUT/AUTHOR/INITIALS/COMMENTS（锚点须是报告原文中能唯一定位的连续文字；脚本已自动去空白+反转义，失配会打印候选段落）。资产评估对"报告"和"说明"各跑一次。
   - `build_crossvalidation.py`：键 SUBTITLE/OUTPUT/TABLE/CHART1/CHART2/CAP1/CAP2/CONSISTENCY/SOURCES(含URL)/CONCLUSION（可选 FONT）。URL 自动做成可点击超链接。
   - JSON 示例（build_annotated）：`{"SRC":"...docx","OUTPUT":"...docx","AUTHOR":"李四","INITIALS":"LS","COMMENTS":[["报告原文连续关键词","批注内容(只描述问题)"]]}`
4. **自检**：编号连续、姓名/项目号正确、无他项目残留、能在 Word 正常打开、链接可点。

## 一致性要点
- 三件套的"形式严格一致"由**这套模板 + 这套脚本**保证。务必复用，勿另起格式。
- 模板在 `../templates/`，脚本用相对路径定位，整个文件夹拷走即可用。
- 覆盖输出时若目标文件正被 Word 打开 → 脚本自动另存"（已修订）"版，请关闭后改用该版或重跑。

## 维护
规则变了改 `../rules/*.md`；格式变了换 `../templates/*.docx` 并同步改脚本里的 XML 片段。建议整个包纳入 git。
