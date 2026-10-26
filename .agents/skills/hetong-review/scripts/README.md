# review_kit.py

纯 Python 标准库（Python 3.9+）、无网络、对合同输入只读。所有写入都只发生在用户显式指定的 `-o` 输出文件，且输出路径不得覆盖输入。

用途：把“原件未变、docx 有哪些易漏内容、数字是否独立复算、报告字段是否齐全”变成可复核记录。
它不做 OCR、不判断法律结论，也不替代对合同全文的人工/模型逐条阅读。

## 1. 审核开始：基线指纹与 docx 覆盖风险

```bash
KIT=<技能目录>/scripts/review_kit.py
python3 "$KIT" fingerprint 原件/合同.docx 原件/附件.docx -o baseline.json
```

输出每份文件的 SHA256、大小和修改时间。docx/dotx 还检测：

- 插入、删除、移动等修订痕迹；
- 页眉页脚、脚注尾注、批注、隐藏文字、文本框、嵌套表格；
- 图片、嵌入对象和 altChunk。

`warnings` 只表示需要额外覆盖，不表示文件无效。普通 `extract` 可能把未接受插入当正文、遗漏删除文字或嵌套表格，因此须先确认签署基线，再决定提取方式。

## 2. 数字候选台账

对 `docx_tool.py extract` 生成的文本、Markdown 或 txt 执行：

```bash
python3 "$KIT" numbers 全文.txt -o numbers.json
python3 "$KIT" numbers 全文.txt --markdown -o numbers.md
```

同时捕捉 ASCII/全角数字、中文数字金额和“百分之……”、财务上下文及 `P0001` / 表格定位。
这是高召回候选清单，会包含与计算无关的条款号和日期；必须逐条筛选语义，不能拿它直接当风险结论。

## 3. Decimal 复算

```bash
python3 "$KIT" calc-tax --net 100000 --rate 6%
python3 "$KIT" calc-tax --gross 113 --rate 0.13
python3 "$KIT" sum --values 113000 10600 --compare 123000
```

`--rate` 用 `6%` 或 `0.06`，不要写裸 `6` 表示 6%。默认按 `ROUND_HALF_UP` 保留两位小数，可用 `--places 0–6` 调整。
税率适用是否合法必须另行核验；算术正确不证明税务处理正确。

复算结果可用 `-o calculation.json` 留存，再把公式、输入、结果和差异写入报告数字底稿。

多分项或多税率合同另建 `ledger.json`：

```json
{"items":[{"id":"设备","net":"100000","rate":"13%","gross":"113000"},{"id":"安装","net":"10000","rate":"6%","tax":"600","gross":"10600"}],"stated_total":"123000"}
```

```bash
python3 "$KIT" tax-check ledger.json -o tax-check.json
```

输出每个分项及 `totals.net` / `totals.tax` / `totals.gross`，并给出与 `stated_total` 的差额。
原文数字不一致记入 `discrepancies`，是审核 finding，不会让计算工具本身失败。报告必须明确写出
三项合计，最终用 `check-report --tax-check tax-check.json` 核对；只算分项或总价差不算完成。

## 4. 审核结束：只读验证

```bash
python3 "$KIT" readonly-check baseline.json 原件/合同.docx 原件/附件.docx -o readonly-result.json
```

传入清单必须与 baseline 相同； changed、missing、added、omitted 或读取错误均返回非零。即使未显式列出文件，输出也不得覆盖 manifest 本身或其中登记的任一原件。
审核阶段只用 `hetong-craft/scripts/docx_tool.py` 的 `extract` / `search` / `verify`，不得使用任何写入子命令。

## 5. 交付物结构门禁

```bash
python3 "$KIT" check-report 审核报告.md -o report-gate.json
python3 "$KIT" check-clauses 保护性条款建议.md -o clause-gate.json

# 最终交付前额外拒绝未勾选项和常见模板占位符
python3 "$KIT" check-report 审核报告.md --final
python3 "$KIT" check-report 审核报告.md --final --tax-check tax-check.json  # 涉价税合同
python3 "$KIT" check-clauses 保护性条款建议.md --final --anchors anchors.json
```

报告检查：必要章节、R-ID 唯一性、每条定位/引文/结论状态/影响/依据/反方复核/修订方向/保护需求及 S1–S4；S1/S2 还须记录第二路径复核。另计算各级 R-ID 数量，并核对“风险分级统计”的数量与名单；最终版缺少该表或仍有占位数量会失败。传入 `--tax-check` 时，还必须在报告中找到合计不含税、合计税额、合计含税三项明确数值。
条款检查：C-ID 唯一性、关联 R-ID、保护目标、完整建议文本、待确认变量、联动和法律核验；最终版必须有“确认方案一致性表”（含确认原句与正文逐字对应）及“逐句反查”记录，防止自检声称与正文不一致。

确认方案中的关键限定词、数字和禁止新增短语可另建锚点：

```json
{"items":[{"c_id":"C-001","required":["完整提交","5 个工作日"],"forbidden":["附带损失","知道或应当知道"]}]}
```

`--anchors` 只逐字检查各 C-ID 的“建议文本”，不检查一致性表里的自我声明。它能发现关键短语丢失或禁词出现，不能理解同义表达、否定句或上下文，因此仍须人工语义复核。

结构通过不代表法律或商业结论正确；仍须执行技能中的实质复核。

## 6. 报告引文与冻结底稿逐字核验

```bash
python3 "$KIT" verify-quotes 审核报告.md 冻结全文.txt 附件全文.txt -o quote-check.json
```

逐条读取 R-ID 下“原文或缺口”字段的块引文，与冻结提取文本做逐字匹配。引文含省略号、找不到原文、没有引文且没有明确缺失声明都会失败；缺失条款声明单列计数，不要求伪造引文。
该命令只证明引文存在，不证明引用上下文完整或法律推理正确。

## 7. 测试

```bash
python3 <技能目录>/scripts/test_review_kit.py -v
```

测试覆盖价税正反算、多分项三项合计与报告门禁、分项差异、非有限数拒绝、docx 修订/页眉/脚注/隐藏文字/文本框/嵌套表格/图片检测、只读验证及默认清单原件防覆盖、数字召回、引文逐字核验、分级统计一致性、确认方案表、逐句反查、确认锚点、禁止覆盖输入和两类模板结构门禁。
