---
name: hetong-craft
description: |
  中文商务合同的起草、修订、润色、完善、审核一体化工作流，含 docx 安全读写、PPTX
  计划式精修、未接受修订痕迹、范围批注、Word 修订/批注时间线深读和无头视觉预览。
  Use when asked to 起草/草拟/撰写合同或条款，修订/修改/变更既有合同文本，润色/语言规范化，
  完善/查漏补缺/补强条款，或审核/审阅/审计合同并出具报告 —— 尤其适用于多主体、多层级、
  背靠背、税率拆分、合同族联动修订等复杂交易结构。统一管理甲乙方买卖条款库、类别多选和未留痕改动检查；原hetong-server的条款与工具已并入本技能。Not for non-hetong legal documents.
---

# 合同工艺（Hetong Craft）

起草、修订、润色、完善、审核中文商务合同的统一工作流。五种模式共用一套铁律与工具，
各自有专门的深度流程文件，按需加载。

## 专用快审子技能

用户点名`hetong-fast-xqy-v1`，或明确要求“xqy快审／徐启源口径快审”时，先转到
[XQY快速审核](../hetong-fast-xqy-v1/SKILL.md)。该分支按八项顺序交付正文宋体的详细Word报告及原稿批注副本，以R编号衔接，
采用其精简受理和报告格式，不先启动下述通用类别问卷或加载深度报告模板。
普通“快速看一下”不自动切换口径；未命中时仍走原五模式流程。

## 真实合同先走受理工具

审核、批注或修订已有合同，在锁定用户指定原件后，先按[受理与宿主交互](references/host-intake.md)取得本次处理口径。技能维护、方法讨论、仅解释一句条文不启动合同问卷。

- 存在 `hetong_intake` 工具时，直接调用它：传源文件、任务产出中的回执路径，以及用户本轮已明确的信息和对应原话。它实际显示选择题并回传结果；不要先写一段文字问卷代替调用。
- Claude Code主代理使用本技能`init → plan --host claude → AskUserQuestion → accept`原生流程，完整参数见[受理与宿主交互](references/host-intake.md)。甲方先选A—D保交付/资金共管组合，再复选其余类别。仅无原生能力的宿主可说明后用`collect`后备；用户要求原生时不擅自换网页。不要把Markdown `[ ]`、单选工具或打印参数当作复选界面。
- 受理核心自动读取类别；合同任务直接使用这些公开命令，不先自行导入`buyer_clause_select`等内部Python函数，不猜测其参数或返回值。函数级调试只用于技能维护。
- 甲乙方买卖条款的预设范围必须有明确类别回执；已有明确文字选择直接记入，不重问。只采用回执中的类别，调用 `export` 取得选中正文。改文前用 `gate --operation comment|revise` 检查回执和原件指纹；回执未完成时可继续提取、盘点和客观核算，不直接写 Word。
- 取消、关闭、超时、仅勾选、宿主缓存或模型自行填写的答案都不构成提交。金额、期限、付款、担保等商业条件仍用普通文字确认，不进入本问卷。交互由主代理完成，不交给子代理猜答案。

## 铁律（任何模式都不得违反）

1. **项目规则优先**。先读取适用的 `AGENTS.md`；`CLAUDE.md` 仅是链接时不重复读取。
   按本次对象加载相关 README、合同规则与交易结构说明，不预读无关项目历史。其中记载的商业实质、
   已锁定的架构决策、红线清单是最高约束：合同文字与商业实质冲突时，改文字、
   不改实质；已否决的方案不得回退。本技能的所有通用清单只是补充，与项目规则
   冲突时以项目规则为准。
2. **合同原件不可擅自改写**。任何 docx/pptx/xlsx 修改：先保全只读基线，默认另存；
   普通工具经明确授权原地修改时先自动备份，红线工具和 PPTX 计划工具不允许原地修改。
   先 `--dry-run`、红线 `plan` 或 PPTX `pptx apply --dry-run` 预演，改后 `verify`
   并重新提取/打开比对（红线分别核对本轮接受/拒绝视图）。审核模式默认只读；
   写入红线稿、批注副本或 PPTX 修订稿前同样须确认基线与输出方式。
3. **各模式的改动权限不得越界**：润色不动实质，修订不夹带未授权的变更，完善
   先提案后落笔。审核与版本对比只读，诊断可含修订方向和保护机制需求；完整条款
   草案与实际修复须在诊断经用户确认后，按修订/完善流程实施，不得夹带商业变更。
   用户明确授权的报告内少量独立建议写法可按专用快审边界提供；不表示条款已获同意或获准写入合同。
4. **可追溯交付**。所有产出附记录：修订出变更日志，审核出审查报告，起草出
   待确认事项清单，完善出建议清单；Office 精修另留输入 SHA、预演/提交结果和
   渲染清单。定位以条款名称/编号为准，不依赖页码。
5. **数字必须复算**。凡金额、税率、比例、期限、违约金、付款节点，一律重新
   计算核对（含税 ÷ (1+税率) 等），并与大小写、前后文、关联合同交叉勾稽；
   不得直接引用示例/模板数字充当合同金额。
6. **保密**。合同与交易结构属敏感信息：不进公开仓库、不外传、不被要求时
   不复述完整交易细节。

## 模式路由

| 用户意图 | 模式 | 加载 |
|---|---|---|
| 从零起草新合同、新条款、补充协议、附件 | 起草 | [mode-draft](references/mode-draft.md) |
| 按明确指令修改既有合同（措辞、条款、数值、结构） | 修订 | [mode-revise](references/mode-revise.md) |
| 只改语言：术语统一、消歧、体例规范、文字打磨 | 润色 | [mode-polish](references/mode-polish.md) |
| 找缺口并补强：缺失条款、风险分配、可执行性 | 完善 | [mode-improve](references/mode-improve.md) |
| 诊断与出报告：合规、逻辑、数字、跨文件一致性、版本对比、复审 | 审核 | [mode-review](references/mode-review.md)（深度流程由 `hetong-review` 承接） |

模式与输出按 [授权边界](references/task-authorization.md) 判断。混合意图而方案未定时先诊断；
已明确的修订范围与有效授权直接沿用，不重复确认。商业条件需要确认时用普通文字等待答复；
仅当歧义实质影响范围或正确性时提问，继续不依赖答案的工作。

**甲乙方条款统一由本技能管理。** 买卖合同需要调用预设条款时，读取[条款库与使用流程](references/trade-clause-guide.md)：甲乙方均先确定本次类别，再按所选范围读取分类Markdown正文。甲方保留A01—A24整合的20项，乙方保留B01—B11；JSON仅为派生索引，不维护另一份正文。条款内容以本技能内的资料为准，`hetong-review`仅引用，不再加载独立的`hetong-server`。
服务器、SN、压测等专用内容仍按实际标的适用，不因移入通用入口而自动套用于服务、租赁等其他合同。两个近似DOCX的额外核对可按[未留痕改动检查](references/hidden-change-review.md)调用本技能的`hidden_change.py`。

新增批注须先按[批注触发标准](references/comment-policy.md)判断：仅命中固定条款且无需改动不批注；实际缺陷、需要增删修订和错别字按本次授权处理。

用户要求履约时间轴，或付款、交付、担保和发票的跨条款关系需要并列展示时，按[合同履约时间轴规范](references/contract-timeline.md)使用固定日期栏＋三栏正文。修订前后复算联动节点，批注只提示具体问题；完整图放分析附件，不自动写入合同。规范含离线Markdown／HTML模板及Word固定列宽要求，不与作者修订活动时间线混淆。

## 通用工作流骨架

1. **受理**：从现有指令提取模式、对象、目标与约束，只追问实质影响结果的缺口。
2. **加载项目规则**（铁律 1）＋ 对象文件全文（docx 用 `extract`）。
3. **定位合同角色**：属于哪个合同族/层级、签约双方、税目税率、与哪些
   文件存在接口（背靠背、验收矩阵、变更规则、附件共用）。
4. **执行模式流程**（见各 mode 文件）。有可独立验收且值得并行的工作包时，按[合同任务协作](references/agent-collaboration.md)派发子代理；模型默认继承主代理，主代理保留整体判断与最终验收。
5. **自检**：模式自检清单 + [审查总清单](references/checklist-review.md)
   中相关部分 + [语言体例](references/style-zh-hetong.md)。
6. **交付**：成果文件 + 对应记录模板（见 [assets](assets/) ），说明未决
   事项与建议下一步。

## 工具速查

**内容已确定且适合委派的 Word 操作**按[Word 操作委派](references/word-operation-agent.md)交接，默认继承主模型及会话运行设置，不固定型号、不限定宿主。不支持子代理或小修改没有分工收益时，由主代理完成。
内容判断和最终验收仍由主 Agent 负责；此入口不扩大写入授权，也不委派尚未确定的条款或批注内容。
写入前按[原生审阅显示与配色](references/native-review-display.md)检查目标客户端的显示能力：仅在能力范围内指定原生颜色，不给正文添加字色、高亮或底纹；不能安全设置时保留原生显示并继续任务。

`scripts/docx_tool.py`（纯 Python 标准库，无需安装依赖）：

```bash
T=<技能目录>/scripts/docx_tool.py
python3 $T extract 合同.docx -o 全文.txt        # 带段落/表格编号提取（P0001/T0001 R01 C01）
python3 $T search 合同.docx 关键词              # 定位段落；--numbers 抓金额税率期限
python3 $T replace 合同.docx 旧 新 --para P0012 -o 合同_修订稿.docx  # 普通清洁稿；不生成留痕
python3 $T batch 合同.docx edits.json --dry-run # 批量修改预演
python3 $T verify 合同.docx                     # 结构校验
```

用户要求 Word/WPS 留痕修改、红线稿或批注时，加载 [红线与批注流程](references/docx-redline.md)，
使用独立的 `scripts/hetong_redline.py`；不要用普通 replace/set 冒充修订，也不要自动接受已有修订。

**已有大量修订/批注的 DOCX，先做 Word 深读。** 在审核、修订或对比之前运行
`scripts/word_reader.py`，生成原文、接受视图、标记视图、时间线、批注线程和条款变化链：

```bash
W=<技能目录>/scripts/word_reader.py
python3 $W analyze 合同.docx --out 产出/word-deep-read
```

必须同时阅读 `original.md`、`accepted.md`、`markup.md`、`deep-read.md` 和 `timeline.csv`，
不能只看某一轮显示文本。时间线中的作者活动时段是推断，不得写成谈判轮次；批注说“已修改”
不等于正文已修复；无锚点、锚点被删除和日期异常必须披露。深读只读，不自动接受/拒绝修订。
复杂结构、历史修订重叠或清洁版派生，按 [工具边界与替代路径](references/word-operation-paths.md) 处理，不假定通用引擎全部支持。
完整方法见 [Word 合同精细阅读](references/word-deep-read.md)，报告骨架见
[Word 深读模板](assets/word-deep-read-report-template.md)。

普通 docx 读写纪律、新合同生成路径（Markdown → HTML → `textutil`（macOS）或
LibreOffice/模板转换（Windows））、编辑操作全集（set/insert-after/delete/add-row）与
失败回退，见 [docx-workflow](references/docx-workflow.md)。

Office 精修、PPTX 计划、无头批量渲染、缓存与视觉验收使用
`scripts/office_kit.py`（自动切换隔离 Python 环境，不依赖 DSH/WPS GUI）：

```bash
O=<技能目录>/scripts/office_kit.py
python3 $O doctor --full                          # 环境、字体和中文文本层自检
python3 $O inspect 合同.docx 汇报.pptx            # 结构盘点、SHA 与 shape_id
python3 $O pptx apply 计划.json --dry-run         # 预演，不写目标
python3 $O pptx apply 计划.json --commit --result 执行记录.json
python3 $O pptx verify 计划.json 产出/修订稿.pptx
python3 $O render 产出/*.docx 产出/*.pptx --out 产出/_预览 --dpi 120
```

多个文件必须合并到一次 `render`，同一文件多处修改合并到一份计划；未变的输入/设置
直接命中缓存。完整编排、PPTX 操作语义、效率与安全边界见
[Office 精修与无头预览](references/office-workflow.md)；计划模板见
[PPTX 计划](assets/pptx-edit-plan-template.json)。不得用普通 XML 改写冒充 PowerPoint
修订痕迹；本工具不生成 PPTX 修订标记或批注。

## 深度参考（按需加载）

- [审查总清单](references/checklist-review.md)：条款维度 + 风险维度 +
  多合同结构专节（分层总分包、三流一致、术语隔离、税率匹配、背靠背传导）
- [中文合同语言与体例](references/style-zh-hetong.md)：润色模式的判定基准
- [Office 精修与无头预览](references/office-workflow.md)：DOCX/PPTX/XLSX/PDF
  的精确修改、渲染、缓存、效率与视觉验收
- [Word 合同精细阅读](references/word-deep-read.md)：修订/批注、时间线、批注线程、
  条款业务变化链和证据分层
- [甲乙方条款库](references/trade-clause-guide.md)：甲方类别多选、A01—A24来源与20项整合条款、乙方B01—B11；按实际交易适用
- [未留痕改动检查](references/hidden-change-review.md)：两份近似DOCX的额外只读核对
- [风险模式库](references/risk-patterns.md)：常见交易结构风险与合同对策
- 模板资产：[统一审查报告](../hetong-review/assets/review-report-template.md) ·
  [变更日志](assets/change-log-template.md) · [起草受理单](assets/draft-intake-template.md) ·
  [Word 深读报告](assets/word-deep-read-report-template.md)
