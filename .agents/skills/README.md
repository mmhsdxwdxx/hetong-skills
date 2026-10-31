# 仓库级合同技能

`hetong-craft`是统一入口和甲乙方条款的唯一内容真源；`hetong-review`承接深度审核；`hetong-fast-xqy-v1`是按XQY八项框架及甲乙方立场交付宋体Word报告的快审子技能。三者共同位于`.agents/skills/`。独立`hetong-server`已撤并，不再维护或加载其副本。

## 目录与职责

```text
.agents/skills/
├── hetong-craft/
│   ├── SKILL.md
│   ├── references/   # 五模式、索引、类别选择、Word与原生审阅显示
│   │   └── clauses/  # buyer/与seller/，一类一个Markdown正文文件
│   ├── scripts/      # Word读写、Office、类别筛选、未留痕检查
│   ├── assets/
│   └── evals/        # 通用与买卖合同虚构场景
├── hetong-review/
│   ├── SKILL.md
│   ├── references/   # 深度审核、对比、关闭核验、法源与保护机制
│   ├── scripts/
│   ├── assets/
│   └── evals/
└── hetong-fast-xqy-v1/
    ├── SKILL.md
    ├── references/   # 八项重点、Word版式与输入规格
    ├── scripts/      # 同源Word/Markdown生成及验证
    ├── assets/       # 纯虚构示例，不含真实交易
    └── evals/        # 行为场景；预期不代表模型实测
```

不按宿主复制实现；支持项目skills发现的宿主从同一目录读取。应从本仓库或其子目录启动/指定工作区。历史调用名称可由craft承接，但不恢复旧独立入口。

## Claude Code CLI 与桌面端

Claude Code 使用项目级 `.claude/skills/` 发现技能。本仓库在该目录为 `hetong-craft`、`hetong-review`、`hetong-fast-xqy-v1` 设置相对软链接，目标仍是本目录的原文件；修改本目录即可更新两端，无须复制技能或上传合同资料。目标项目的AGENTS.md／CLAUDE.md规则由该项目自行维护，本公开包不附带真实合同工作区规则。

- CLI：在本仓库或其子目录运行 `claude`，使用 `/hetong-craft` 或 `/hetong-review`，也可直接说明合同任务。
- 桌面端：在 Claude 的 **Code** 页选择本仓库作为本地项目，新建会话后输入同样的命令。此连接适用于 Code 本地会话，不等同于 Chat、Cowork 或云端账号技能安装。
- 原生子代理：项目提供 `hetong-worker`，其定义软链接到 `hetong-craft/agents/claude-hetong-worker.md`。显式 `model: inherit` 并预载 `hetong-craft`，按[协作规则](hetong-craft/references/agent-collaboration.md)交接，避免全局子代理默认模型替代主模型。未设置提权、自动放行或全局模型变更。
- 工作树：CLI 2.1.277 及以上在工作树根目录没有 `.claude/skills` 时可读取主检出的技能；已有自己的目录时只读工作树内那一份，不会自动追踪主检出的后续修改。当前桌面 Code 启动的工作树从主检出读取技能等项目配置，`CLAUDE.md` 则仍取工作树版本。不要在 CLI 工作树留下不完整的同名目录；本地直开本仓库可直接使用当前开发版。

软链接只做发现适配；正式分发时必须同时包含链接及 `.agents/skills/` 目标。在不支持软链接的环境应使用宿主支持的安装方式，不能把链接文本当作技能内容。规则导入使用普通文本，兼容不能创建软链接的系统。

发现机制、规则导入和桌面端范围分别按 [Skills](https://code.claude.com/docs/en/skills)、[项目规则](https://code.claude.com/docs/en/memory#share-one-file-with-other-coding-tools)、[Desktop](https://code.claude.com/docs/en/desktop#use-skills) 与 [工作树](https://code.claude.com/docs/en/worktrees#what-worktrees-share-with-the-main-checkout) 官方说明核对；宿主加载能力须以实际版本探针为准。

## 调用与条款选择

真实合同先执行[受理工具与宿主交互](hetong-craft/references/host-intake.md)。PiDeck项目扩展为`hetong_intake`/`hetong_clause_categories`（保留`hetong_buyer_categories`兼容名）提供原生问答和多选，`.pi/extensions/hetong-intake.ts`仅链接到craft中的适配器；不依赖PiDeck全局内置问答扩展是否启用。Claude桌面Code和CLI由主代理使用`plan --host claude → AskUserQuestion → accept`原生交互，也可调用项目命令`/hetong-intake`。甲方先选A—D保交付/资金共管组合，再复选必要、一般、其他（附“其余类别均不选”）；不因五类数量自动转网页。两者共用问题、类别数据、回执及权限检查，不按模型品牌分支；DeepSeek-V4.1-Flash的调用行为须单独实测。

PiDeck需在本仓库根目录启动新会话并允许项目扩展；从未加载项目扩展的子目录启动时走本机表单后备，不宣称已弹原生界面。已有会话的重载仍按用户要求，不自动刷新。

- `请用hetong-craft审核这份采购合同，我方是买方，先只审不改。`
- `我方为甲方，本次预设条款选保交条款和必要条款，按已确认条件另存修订批注版。`
- `请用hetong-craft审核销售合同，我方是乙方，核对B01—B11。`
- `请对比这两版合同，并检查有没有没有修订记录的直接改文。`
- `xqy快审，合同位置：……。`（或显式`$hetong-fast-xqy-v1`／Claude的`/hetong-fast-xqy-v1`）
- `xqy快审，站乙方立场，合同位置：……。`（在对应论述中突出本方固定条款）

XQY快审保留八项顺序，按本次甲方／乙方及实际买卖角色切换方向，并突出本方固定条款。默认交付两份Word：11pt宋体的详细审核报告，直接给出文字修订建议；保留原稿正文及历史的原生批注副本，简述注意或修订点。两份共用R编号，批注指向报告的对应详细项。遵守已有类别选择／排除，不弹整套问卷、不自动改写正文。普通审核仍按原流程执行。详见[快审子技能](hetong-fast-xqy-v1/SKILL.md)。

甲方按[类别选择](hetong-craft/references/buyer-category-selection.md)多选，乙方通过同一受理核心选择本方现有类别，再读取选中条款。当前[条款目录](hetong-craft/references/buyer-clause-catalog.json)把A01—A24来源整合为20项，按保交、必要、一般、其他、共管分类；不默认全选，不自动带入未选类别。乙方B01—B11及说明也在craft中，按用户逐条改稿分为必要10项、其他1项。内容以craft为准，review不复制第二份。

[买卖条款使用指南](hetong-craft/references/trade-clause-guide.md)负责角色和适用边界。目录归并不把服务器品名、SN、压测等要求泛化为所有合同义务。类别选择不是商业条件确认，也不是追加或删除条款的授权。

[按需协作](hetong-craft/references/agent-collaboration.md)默认继承主模型。[原生审阅显示](hetong-craft/references/native-review-display.md)只在Word/WPS能力范围内指定；不能安全设置时保留原样并继续，不向正文加字色、底纹，也不改变历史作者。

## 履约时间轴共用规范

用户要求时间轴，或需要看清付款、交付、担保及开票的联动时，三个入口均引用craft的[合同履约时间轴规范](hetong-craft/references/contract-timeline.md)，不另建重复子技能。固定日期栏＋三栏正文，默认A4版心159.2mm，四列22／52／50／35.2mm，末栏“发票（万元）”，无合并标题行，每要点独立换行；区分应付、已付、待退和资金占用。

[Markdown虚构示例](hetong-craft/assets/contract-timeline-example.md)可独立阅读；[HTML模板](hetong-craft/assets/contract-timeline-template.html)内嵌样式、离线打开及打印，Word采用固定列宽表格。完整图放分析报告或获准附件，批注只写具体问题；示例数值不能成为商业默认值。此为履约事件图，不是`word_reader.py`的作者修订活动时间线。

## 维护与验证

保持相对引用有效，真实合同和任务成果不得混入技能包。Office环境放在用户数据或任务产出目录，技能中不放缓存、锁文件或虚拟环境。每次改动后运行：

```bash
python3 .agents/skills/hetong-craft/scripts/test_hetong_redline.py
python3 .agents/skills/hetong-craft/scripts/test_word_reader.py
python3 .agents/skills/hetong-craft/scripts/test_office_kit.py
python3 .agents/skills/hetong-craft/scripts/test_platform_compat.py
python3 .agents/skills/hetong-craft/scripts/test_review_audit.py
python3 .agents/skills/hetong-craft/scripts/test_hidden_change.py
python3 .agents/skills/hetong-craft/scripts/test_buyer_clause_select.py
python3 .agents/skills/hetong-craft/scripts/test_clause_library.py
python3 .agents/skills/hetong-craft/scripts/test_hetong_intake.py
python3 .agents/skills/hetong-craft/scripts/test_intake_redline.py
python3 .agents/skills/hetong-review/scripts/test_review_kit.py
"$PY" .agents/skills/hetong-fast-xqy-v1/scripts/test_build_report.py
"$PY" .agents/skills/hetong-fast-xqy-v1/scripts/test_build_bundle.py
```

快审生成测试中的`PY`使用现有Office Python（需要python-docx）。涉及路由、目录或frontmatter时，另做宿主发现与加载探针，确认发现craft、review和fast-xqy-v1三个入口且引用可读，无旧server入口。更新后的核验使用独立新进程；是否刷新已有对话按用户指令，不把发送刷新消息当成加载成功。

涉及受理交互时，加载探针不构成验收：另验证PiDeck/Claude桌面界面、目标模型实际调用、两项以上同时选择及回传、取消/空选/超时和禁止越权写入。记录软件版本、模型实际标识、真实界面与模拟测试的区别。

Office环境变化后运行`hetong-craft/scripts/office_kit.py doctor --full`及相应回归。显示与实机兼容必须据实际证据报告，脚本通过不等于Word/WPS界面已验证。

开发和发布规则见[DEVELOPMENT.md](DEVELOPMENT.md)。历史验证不替代当前验证，当前记录见[VALIDATION.md](VALIDATION.md)。

行为场景分别读取`hetong-craft/evals/evals.json`、`hetong-craft/evals/trade-evals.json`、`hetong-craft/evals/intake-evals.json`和`hetong-craft/evals/comment-evals.json`，不能只运行第一份。场景中的`files`均相对各自技能根目录解析；行为评测记录与单元回归结果分开，不能将预期答案当实际通过。

## 正文与批注维护

甲乙方条款均在`hetong-craft/references/clauses/`按一类一个Markdown维护；JSON只存索引，旧整合预览和乙方基线文件仅保留入口。正文调整不要求改SKILL.md；改名、增删或移动条款后执行`buyer_clause_select.py --side seller reindex`预览（甲方使用`--side buyer`），核对后加`--write`更新索引，再运行`validate`。完整目录必须随技能一并打包，不能仅分发SKILL.md或JSON索引。

新增批注执行[统一触发标准](hetong-craft/references/comment-policy.md)：已符合固定条款且无需改动不批注；真实问题、必要的增删修订和错字按授权处理。内部覆盖记录与Word批注分开；本轮修复说明和历史批注继续保留。
