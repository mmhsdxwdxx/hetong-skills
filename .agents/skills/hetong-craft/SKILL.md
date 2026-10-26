---
name: hetong-craft
description: |
  中文商务合同的起草、修订、润色、完善、审核一体化工作流，含 docx 安全读写、PPTX
  计划式精修、未接受修订痕迹、范围批注、Word 修订/批注时间线深读和无头视觉预览。
  Use when asked to 起草/草拟/撰写合同或条款，修订/修改/变更既有合同文本，润色/语言规范化，
  完善/查漏补缺/补强条款，或审核/审阅/审计合同并出具报告 —— 尤其适用于多主体、多层级、
  背靠背、税率拆分、合同族联动修订等复杂交易结构。Not for non-hetong legal documents.
---

# 合同工艺（Hetong Craft）

起草、修订、润色、完善、审核中文商务合同的统一工作流。五种模式共用一套铁律与工具，
各自有专门的深度流程文件，按需加载。

## 铁律（任何模式都不得违反）

1. **项目规则优先**。动手前必须先在当前仓库（含祖先目录）查找并通读 `AGENTS.md`、
   `CLAUDE.md`、`README` 及任何合同规则/交易结构说明文档。其中记载的商业实质、
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

意图混合时（如"审一遍顺便改好"）：先按审核模式出诊断，经用户确认后再切
修订/完善模式实施；两步的产出记录分开。意图不明时用 `ask_user_question`
确认模式与范围，不要自行假设。

**服务器采购/销售合同**：在通用模式之上叠加 [hetong-server](../hetong-server/SKILL.md)
专用 profile（十一项商业底线、交付/验收/担保/现金流）；该 profile 只适用于服务器设备
采购/销售，不把其中底线泛化到其他合同。加载后仍以本技能和 `hetong-review` 的权限、
R/Q/C 追溯、Word 深读和交付门禁为准。

## 通用工作流骨架

1. **受理**：明确模式、对象文件、用户目标与约束；信息不足先问。
2. **加载项目规则**（铁律 1）＋ 对象文件全文（docx 用 `extract`）。
3. **定位合同角色**：属于哪个合同族/层级、签约双方、税目税率、与哪些
   文件存在接口（背靠背、验收矩阵、变更规则、附件共用）。
4. **执行模式流程**（见各 mode 文件）。
5. **自检**：模式自检清单 + [审查总清单](references/checklist-review.md)
   中相关部分 + [语言体例](references/style-zh-hetong.md)。
6. **交付**：成果文件 + 对应记录模板（见 [assets](assets/) ），说明未决
   事项与建议下一步。

## 工具速查

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
完整方法见 [Word 合同精细阅读](references/word-deep-read.md)，报告骨架见
[Word 深读模板](assets/word-deep-read-report-template.md)。

普通 docx 读写纪律、新合同生成路径（Markdown → HTML → textutil 转 docx）、
编辑操作全集（set/insert-after/delete/add-row）与失败回退，见
[docx-workflow](references/docx-workflow.md)。

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
- [服务器合同专用 profile](../hetong-server/SKILL.md)：十一项商业底线、服务器交付/验收/
  担保/现金流分析；只在服务器采购/销售任务加载
- [风险模式库](references/risk-patterns.md)：常见交易结构风险与合同对策
- 模板资产：[审查报告](assets/review-report-template.md) ·
  [变更日志](assets/change-log-template.md) · [起草受理单](assets/draft-intake-template.md) ·
  [Word 深读报告](assets/word-deep-read-report-template.md)
