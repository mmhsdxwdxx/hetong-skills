# 仓库级合同技能

本目录是 `hetong-craft`、`hetong-review` 与 `hetong-server` 的**唯一运行时真源**，供 Pi、Codex、DSH 在 `contact` 仓库中共同使用。不要再从其他目录复制或同步第二份实现。

## 目录

```text
.agents/skills/
├── hetong-craft/       # 起草、修订、润色、完善、审核统一入口；DOCX/PPTX 精修与无头预览
│   ├── SKILL.md
│   ├── references/
│   ├── assets/
│   └── scripts/
├── hetong-review/      # 深度审核、跨合同勾稽、版本对比、关闭核验与保护条款
│   ├── SKILL.md
│   ├── references/
│   ├── assets/
│   ├── scripts/
│   └── evals/          # 仅虚构评测材料
└── hetong-server/      # 服务器采购/销售合同专用 profile；乙方 B01—B11、甲方 A01—A04
    ├── SKILL.md
    ├── references/     # 两套立场基线、量化、受理、修订、批注、暗改检测
    ├── scripts/        # hidden_change.py（两版近似合同未留痕改动）
    └── agents/         # 可选宿主展示元数据
```

三个宿主都原生发现项目 `.agents/skills/`：Pi 与 Codex 直接使用，DSH 以 `project-agents` 来源加载。无需 `.pi`、`.dsh` 副本、软链接或额外适配器。必须从本仓库（或其子目录）启动/指定工作区；在其他项目目录中不会加载本仓库技能。

## 调用

自然语言即可：

- `请使用 hetong-craft 起草一份技术服务合同，我方是委托方。`
- `请审核这份合同，先出诊断，不修改原件。`
- `按已确认方案另存 Word 红线稿，保留已有修订并添加批注。`
- `这份服务器采购合同修订和批注很多，请先重建原文/接受视图和时间线，再按服务器底线审核。`
- `按计划精修这份 PPTX，先预演；提交后批量渲染并检查中文和版式。`
- `请使用 hetong-server 审核服务器销售合同，我方是卖方，只审不改。`
- `请使用 hetong-server 审核服务器采购合同，我方是甲方，重点核对 A01—A04 保函与共管账户。`
- `这是同一合同的前后两版，请正常对比之余，检查有没有绕过修订直接改正文的地方。`

Pi 也可显式使用 `/skill:hetong-craft`、`/skill:hetong-review` 或 `/skill:hetong-server`。其他宿主按各自的技能选择方式调用，或直接说明技能名。修改技能后，新开会话或执行宿主的重载操作；Pi 使用 `/reload`。

## 维护规则

1. 只修改本目录中的实现；技能之间的相对路径必须保持有效。
2. 合同文本核心脚本以 Python 3.9+ 标准库为主；Office 精修/无头渲染通过 `scripts/office_kit.py` 自动切换到隔离运行环境，依赖锁定在 `scripts/office-requirements.lock`。LibreOffice 只用于转换与渲染，不依赖 DSH、WPS 服务或上传合同。
3. 真实合同、审核报告、用户交易信息和临时输出不得放入技能目录；仍按仓库 `AGENTS.md` 的任务目录与保密规则处理。经用户明确授权的领域基线（如 `hetong-server` 的 B01—B11 与 A01—A04）可作为技能资料保留，但不得附带真实交易材料。
4. 不把 `.pyc`、`__pycache__`、Office 锁文件或测试临时文件提交入库。
5. 每次改动后运行：

```bash
python3 .agents/skills/hetong-craft/scripts/test_hetong_redline.py
python3 .agents/skills/hetong-craft/scripts/test_word_reader.py
python3 .agents/skills/hetong-craft/scripts/test_office_kit.py
python3 .agents/skills/hetong-craft/scripts/test_platform_compat.py
python3 .agents/skills/hetong-craft/scripts/test_review_audit.py
python3 .agents/skills/hetong-server/scripts/test_hidden_change.py
python3 .agents/skills/hetong-review/scripts/test_review_kit.py
```

Office 环境或字体变化后先运行：

```bash
python3 .agents/skills/hetong-craft/scripts/office_kit.py doctor --full
```

使用与排障见 [hetong-craft Office 流程](hetong-craft/references/office-workflow.md)。Word 修订/批注深读见 [Word 合同精细阅读](hetong-craft/references/word-deep-read.md)。服务器合同专用方法见 [hetong-server](hetong-server/SKILL.md)，其中乙方 B01—B11 与甲方 A01—A04 只适用于对应立场；两版近似合同的暗改检测见 [未留痕改动说明](hetong-server/references/hidden-change-review.md)。

6. 涉及路由、frontmatter 或目录结构时，另做 Pi、Codex、DSH 发现与加载探针；不能只凭文件存在判断可用。

本次迁移与宿主验证结果见 [VALIDATION.md](VALIDATION.md)。
