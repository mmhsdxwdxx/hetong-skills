# 受理工具与宿主交互

适用真实已有合同的审核、批注、修订；不在技能维护或一般讨论中启动。先锁定用户指定源文件，已有明确答复直接沿用。界面只收集工作范围，不收集或批准重要商业条件。

## 统一状态与数据

`scripts/hetong_intake.py`是问题、分支、回执和条款范围的单一实现。甲方与乙方分别绑定本方索引及分类Markdown内容的指纹。立场、合同阶段、处理方式及交付形式按已有答案裁剪；甲乙方类别分别读取当前`buyer-clause-catalog.json`和`seller-clause-catalog.json`，没有第二份静态类别表。署名未指定时沿用已有SR规则，不强制多问一轮。

真实合同任务只使用本页公开CLI或专用工具。无需先另读类别、导入Python函数或自行组装问卷；不要猜测`load_catalog`等函数的参数及返回结构。维护者调试内部实现与合同受理是两条路径。

回执绑定源文件绝对路径/SHA、目录SHA、任务ID、问题批次和实际答案。`pending`或`cancelled`均未完成；只有经过完整提交或本次用户明确文字指令形成的`submitted`状态可用于后续写入。切换合同、修改源文件或修改目录后不得复用旧回执。不得把模型猜测写入`known`，也不得伪造原生工具结果。

## PiDeck

本仓库项目扩展注册`hetong_intake`、`hetong_clause_categories`，保留`hetong_buyer_categories`兼容入口。前者完成缺失的受理项和必要类别选择；类别工具适用于立场已经确定、只补相应类别的情况。调用时提供：

- `receipt_path`：本任务`产出/`中的JSON回执绝对路径；
- `sources`：本次原件绝对路径数组；已有有效回执时可复用；
- `known`与`known_evidence`：仅用户已经明确的信息及其原话，例如`{"perspective":"buyer","action":"review"}`。未知项不填。

扩展使用PiDeck原生批量问答桥接，`review:true`保留最终提交步骤。类别分为下述前置单选和其余类别`multi_select`，不能将复选改成单选或自由文本题。原生界面勾选中的草稿不保存为完成回执。

该扩展是本仓库的专用入口，不要求恢复PiDeck全局`ask_question`扩展，也不引入其“所有问题都用卡片”的规则。缺少项目扩展、从未加载它的子目录启动、普通Pi或RPC不支持该桥接时，走下述本机表单；不得宣称已显示原生复选框。

## 甲方类别：前置单选＋其余类别复选

前置题是“本次纳入哪组保交付、资金共管条款？（仅选择条款范围）”，按用户指定顺序显示：

- A、保交付 + 不资金共管 → 保交条款
- B、不保交付 + 不资金共管 → 两类均不纳入
- C、保交付 + 资金共管 → 保交条款、共管条款
- D、不保交付 + 资金共管 → 共管条款

随后复选：必要条款、一般条款、其他条款、其余类别均不选。前三项可并选，第四项与前三项互斥。选B且选“其余类别均不选”才表示本次不调用预设条款。核心把两题合并成完整`categories`回执；任一题缺失均不提交。目录内容和类别标识仍从原数据源读取，不改条款正文。

“保交付”“资金共管”仅用于选择相应预设条款，不确认履约承诺、资金安排、金额或担保条件，也不删除原合同中未选类别的已有约定。

## 乙方类别

乙方只显示当前有条目的类别：必要条款、其他条款，以及互斥的“本次不用乙方预设条款”，使用原生多选。两个类别可并选；不自动全选，不套用甲方保交付/资金共管组合题。使用同一个`categories`回执字段，切换甲乙方时清除旧方选择，重新绑定本方库。

## Claude 桌面 Code 与 CLI：优先原生交互

主代理使用Claude原生`AskUserQuestion`。每次最多四题、每题二至四项；类别两题分别为`multiSelect:false`和`multiSelect:true`。该布局符合[官方工具格式](https://code.claude.com/docs/en/agent-sdk/user-input#question-format)，不能因模型品牌改走网页，也不能让子代理代问。

以技能实际路径代入`I`，在任务产出目录执行：

```bash
I=<技能实际目录>/scripts/hetong_intake.py
python3 "$I" init 产出/受理回执.json --source /绝对路径/合同.docx
python3 "$I" plan 产出/受理回执.json --host claude --out 产出/问题批次.json
```

计划`questions`为空时直接读取`status`返回，不调用空问卷。否则读取计划，把`arguments`原样传给实际`AskUserQuestion`，不要自行改写问题、预填答案或仅打印参数。收到真实工具结果后，原样保存为`产出/原生答复.json`，再执行：

```bash
python3 "$I" accept 产出/受理回执.json --plan 产出/问题批次.json --response 产出/原生答复.json --host claude
```

若状态仍为`pending`且确有下一批问题，继续`plan → AskUserQuestion → accept`；取消、跳过、超时或自由文本改意时停止自动弹问，不替用户补答案。多选答案允许原生返回的标签数组或逗号分隔标签，核心按所展示选项逐一核验。类别题不会自动转网页；未来类别增加导致超限时明确报错，须调整布局。

用户已有明确答复时先存为JSON，用`init --known 已确认项.json --evidence '对应的用户原话'`记录，避免重问。键和值：`perspective=buyer|seller|other`、`stage=unsigned|signed|performing|unknown`、`action=review|comment|revise`、`delivery=tracked|clean|both`、`author`及显式`categories`数组。`categories:[]`仅在用户明确不用预设库时填写，甲乙方均适用。不要把界面前置题当作具体商业条件确认。

## 无原生交互能力的后备

仅当当前宿主确实没有可用原生工具时，说明能力缺口后用`collect`打开127.0.0.1本机表单；这是后备，不能记作Claude原生复选验收。用户要求原生而宿主不具备时，保留待选择，不能擅自换网页。

```bash
python3 "$I" collect 产出/受理回执.json
```

`collect`显示同样两题，打印本机链接并等待提交；只缺类别时可加`--phase categories`。不要反复新开表单。关闭页面、超时不套默认值。表单不上传源文件，也不接收自定义商业问题。远程会话无法访问时说明限制并保留待选择。

所有宿主的原生结果出现`afkTimeoutMs`、`followUp`、未答注记、取消或答案缺失时均不得按成功处理。校验拒绝未知类别、空而未明示的选择、互斥项并选、重复/过期提交及源文件漂移。能力识别看实际工具与宿主，不按模型品牌判断。

## 使用回执

```bash
python3 "$I" status 产出/受理回执.json
python3 "$I" export 产出/受理回执.json --out 产出/选中甲方条款.md
python3 "$I" gate 产出/受理回执.json --operation revise
```

`export`只导出明确选中类别。`gate`同时核对处理方式和原件指纹：只审不改不能通过修订检查，只加批注不能取得正文修订权限。真实合同调用红线引擎`apply`时附`--intake-state 产出/受理回执.json`，由引擎在写入前再次核对。已签或状态未明的修改仍是拟变更建议，不声称编辑文件改变合同效力。具体金额、期限、担保及旧修订取舍按原有授权规则另行处理。
