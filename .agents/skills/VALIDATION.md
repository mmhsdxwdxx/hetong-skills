# 仓库级迁移与验证记录

## 迁移结果

- 运行时真源：`contact/.agents/skills/hetong-craft/`、`hetong-review/` 与 `hetong-server/`。
- 共迁移 39 个技能文件（17＋22），保留 `references`、`assets`、`scripts`、`evals` 和相对路由；后续并入服务器 profile 与审阅审计器。
- 为兼容本机默认 Python 3.9，`hetong_redline.py` 增加仅类型检查使用的 `TypeGuard` 回退；算法未改变，71 项测试在 3.9 下通过。
- 移除旧 Pi 用户目录绝对路径；红线说明改为宿主无关表述。
- 未创建 `.pi`、`.dsh` 适配副本或软链接；三个宿主共用同一目录。
- 原 Pi 全局目录已移至技能扫描范围外的备份：`$HOME/.pi/agent/workspaces/contact-repo-skill-migration/pi-global-backup/`。其他全局技能未改动。
- 未读取、复制或修改真实合同；未改动 DSH、Codex 全局配置。

## 两版近似合同的未留痕改动

`hetong-server/scripts/hidden_change.py` 检测两版 DOCX 的未留痕文字改动，并用 `HC` 编号输出报告。
它以“后一版拒绝全部修订”与前一版基线比较，区分直接改文、可由修订解释的改动、审阅历史减少和非文本变化。

使用虚构样本验证：直接改文、真实留痕改文、混合、净额修订历史清除、批注清除、隐藏文字、媒体替换、低相似度和输出覆盖保护。
工具只读，不自动接受/拒绝修订；结论不指控主观恶意，仍须核对来源、授权和沟通记录。


已把 Codex 导出的 `hetong-server` 作为**服务器采购/销售合同专用 overlay** 并入：

- `hetong-server/SKILL.md`：专用触发、乙方 B01—B11 与甲方 A01—A04 对照、修订/批注规则；
- `references/`：开场分流、量化与履约、修订、批注、乙方十一项基线、甲方 A01—A04；
- 复用 `hetong-craft` 的 `word_reader.py`、`hetong_redline.py`、`office_kit.py`；
- 新增甲方视角 `buyer-protection-clauses.md`（A01—A04），与乙方 B01—B11 分列，避免两套立场互相替代；
- 将独立审阅结构审计器纳入 `hetong-craft/scripts/audit_docx_review.py`，
  测试纳入 `test_review_audit.py`；
- `agents/openai.yaml` 仅保留为可选宿主展示元数据，不构成第二份实现；
- `B01—B11` 与 `A01—A04` 只在服务器 profile 中按立场启用，不替换通用审核的 R/Q/C、S-ID 或只读闸门。

这套 profile 的十一项条文来自用户提供的商业基线，不等同于法律有效性结论；其中的历史台数、比例、金额只作上下文，不无条件套用到其他合同。导入源包 SHA-256 为 `94536649af44ae8808a4a921aad75b9da0daa21f837d5b1feca9dd6badccadd4`（10 个条目），未包含真实合同或历史任务；`INSTALL.md` 未作为运行时入口复制，改为本仓库的集成说明。

## 名称迁移

运行时目录与技能名已统一为 `hetong-craft`、`hetong-review`；红线脚本改为
`hetong_redline.py`，对应测试改为 `test_hetong_redline.py`，语言体例文件改为
`style-zh-hetong.md`。当前技能树、路由、文档和 `AGENTS.md` 已无旧英文前缀。
本次通过改名前的完整备份逐文件比对确认：只发生了名称替换，没有改写合同方法、法律内容、
业务逻辑或脚本算法；改名后全部回归和路由探针重新通过。

## 自动回归

| 检查 | 结果 |
|---|---|
| `test_hetong_redline.py` | 71 项通过 |
| `test_review_kit.py` | 12 项通过 |
| `test_office_kit.py` | 12 项通过（环境、CJK 渲染、缓存、安全检查、PPTX 全流程） |
| `test_review_audit.py` | 13 项通过（成品结构、历史批注/修订保留、锚点、关系、只读、ZIP/DTD 安全） |
| `test_platform_compat.py` | 5 项通过（Windows venv、LOCALAPPDATA、LibreOffice 路径、系统字体、导入不依赖 Unix 模块） |
| `test_hidden_change.py` | 8 项通过（直接改文、留痕改文、混合、历史清除、批注清除、隐藏文字/媒体、低相似度、覆盖保护） |
| `test_word_reader.py` | 5 项通过（16 条修订、6 条批注、移动/格式/表格/隐藏文字/异常时间线/净额回改） |
| Ruby/Psych 实际解析两份 frontmatter | 通过；目录名与 `name` 一致，description 合规 |
| 77 个相对 Markdown 链接、60 个技能文件、必要资源、无软链接、无 Python/Ruff 缓存、无用户目录绝对路径 | 通过 |
| 活跃全局同名技能 | 0；Pi 不会与旧副本冲突 |

迁移过程只对仓库版测试文件做了机械 lint 清理及一处无用变量删除；71 项测试在修改后重新通过。生产引擎未因本次迁移改变算法。

5 个核心脚本完成 LSP 检查；另有 9 条 ast-grep“潜在未捕获异常”辅助提示，均已核对为迁移前代码的错误处理风格。红线引擎和审核工具有最外层异常处理，普通 docx 工具保持原行为；本次不为路径迁移做无关生产改写。

## Word 修订/批注深读

`hetong-craft/scripts/word_reader.py` 是只读标准库解析器，不依赖 Office 运行环境、Word/WPS
或上传服务。它从 OOXML 生成 `original.md`、`accepted.md`、`markup.md`、`timeline.csv`、
`comments.csv`、`deep-read.json` 和事实层报告。

回归覆盖：

- 原文/接受视图：插入、删除、移动、段落标记插入/删除和表格行插入/删除；
- 传统批注与 `commentsExtended` 回复/解决状态；批注锚点删除、无锚点正文、有锚点无正文；
- 作者日期排序、未来日期、缺失日期、时区未标明；
- 字符格式修订、隐藏文字（直接属性和段落样式）、条款分组、业务关键词信号；
- 输出文件完整性、输入 SHA 不变、禁止覆盖未授权输出。
- 另用一套包含 17 条修订、6 条批注的完整虚构 DOCX 做交互样例：`office_kit.py inspect`
  能重新读取，LibreOffice 无头渲染生成 1 页页面图/联系表；页面图中的中文、插入/删除和表格可见。
  PDF 文本层会同时包含修订前后的文字，不能作为阅读正确性的唯一证据；本流程以 OOXML 视图为准。

测试使用现场构造的虚构 OOXML（不是真实合同，也不是 Word 导出图片），不把“能解析”说成
“已通过 Word/WPS 显示兼容”。实际使用仍应在 Word/WPS 核对自动编号、复杂表格、域和最终分页。


### 组件

| 组件 | 版本 / 位置 | 用途 |
|---|---|---|
| LibreOffice | 26.8.1.1，`/Applications/LibreOffice.app` | 批量无头 DOCX/PPTX/XLSX → PDF |
| 隔离 Python | 3.12，`$HOME/.local/share/contact-office/venv` | Office 结构与精修，不污染系统 Python |
| python-docx / python-pptx | 1.2.0 / 1.0.2 | DOCX 结构、PPTX 精修 |
| openpyxl / lxml / defusedxml | 3.1.5 / 6.1.3 / 0.7.1 | XLSX 检查、OOXML 解析 |
| Pillow / pypdf / pypdfium2 | 13.0.0 / 6.21.2 / 5.14.0 | 页面图、PDF 元数据和渲染 |
| Noto Sans/Serif CJK SC | 用户字体目录，常规/粗体 | LibreOffice 中文回退字形 |

Python 版本由 `scripts/office-requirements.lock` 固定。字体包来自 Debian `stable` 镜像的
`fonts-noto-cjk 20240730+repack1-1`，安装前核对了镜像索引中的 SHA-256
`f5dc28a754e17327d99f0a612134d92c8dd6187314ae967cb77f25df60860139`，下载文件一致。

Office 工具自动生成专用 fontconfig 配置，只通过环境变量作用于 LibreOffice 子进程；
未修改用户全局字体配置或 LibreOffice 应用包。

### 结果

- `doctor --full` 通过：LibreOffice 可执行、Python 库齐全、PDFium 可生成页面图，
  DOCX/PPTX 两份虚构文件一次渲染成功，且 PDF 中文文本层可提取（唯一汉字数 31）。
- PPTX 精修闭环通过：dry-run 不写目标；commit 后源文件 SHA 不变、媒体不变、
  未指向形状不变；表格、备注、字体、几何与文本修改均从重开文件复核；
  跨 run、样式合并、越界、SHA 不符和原地覆盖均被拒绝或要求显式授权。
- 渲染缓存通过：相同输入/SHA/DPI/字体命中缓存且不启动 LibreOffice；篡改页面图后
  自动失效重渲染。宏文档、ZIP 路径穿越被拒绝。
- 视觉检查：DOCX/PPTX 联系表和原始页图中文均正常显示，不是占位字；金额、表格、
  修改后的 40% 和含税标注可见。
- Pi 新进程实测加载了更新后的 `hetong-craft`、`office-workflow.md` 和 PPTX 计划模板，
  `office_kit` 路由与子命令识别正确；Python LSP 对新 Office 脚本 0 诊断，Ruff 关键错误检查通过。
- Word 深读能力另做 Pi 探针：从 `hetong-review` 入口读取 `word-deep-read.md` 和
  `word_reader.py`，确认三视图/时间线/只读命令路由可加载；Python LSP 0 诊断，Ruff 关键错误检查通过。

### 效率实测

虚构 6 个文件、共 9 页，同一台机器：

| 模式 | LibreOffice 启动 | 墙钟时间 |
|---|---:|---:|
| 逐文件 6 次 | 6 | 15.293 秒 |
| 单批 1 次 | 1 | 4.079 秒（约 3.75 倍） |
| 相同输入缓存复用 | 0 | 0.389 秒 |

### 未完成的目标软件验证

- QuickLook 成功生成合成的 DOCX 缩略图，中文可见；但 PPTX 预览超过 120 秒未完成，
  已终止本次预览进程。
- 因此当前可确认 **LibreOffice 无头渲染、DOCX/PPTX 结构语义和 python-pptx 重开验证**，
  不能宣称 Microsoft PowerPoint、WPS 的最终显示与动画兼容性已经通过。
- 关键交付仍需目标软件人工打开一次；首版不支持 PPTX 修订标记/批注，也不支持
  SmartArt、图表内部对象和幻灯片增删。

## 宿主实测

| 宿主 | 实测范围 | 结果 |
|---|---|---|
| Pi | 新进程启动；显式加载 `hetong-craft → mode-review → hetong-review` | 通过；主链两个 `SKILL.md` 均来自本仓库，五模式正确 |
| Pi（服务器 profile） | 新进程启动；显式加载 `hetong-server → eleven-baselines → buyer-protection-clauses` | 通过；主链来自本仓库，乙方/甲方基线与专用触发读取正确 |
| Codex CLI 0.163.0 | 只读、临时会话；加载 `hetong-server` 与对应立场基线，运行 `audit_docx_review.py --help` | 通过；文件读取和脚本退出状态 0（Firecrawl MCP 认证错误与该探针无关） |
| DSH 0.2.1-rc.2 | 直接调用本机 `FileSystemSkillProvider`，隔离用户技能根 | 通过；两份技能均为 `project-agents`，正文与资源基目录加载正确，目录观测完整 |

这些探针证明**项目发现、显式加载、相对资源和脚本入口**可用，不是全部自然语言触发率或完整合同任务的质量基准。Python 3.9 本机直接运行 `test_hetong_redline.py`、`test_word_reader.py`、`test_review_audit.py`、`test_review_kit.py` 和 `test_office_kit.py` 均已通过；Office 测试由隔离环境执行。

## DSH 宿主环境的独立阻塞

DSH 全量 headless 启动尚未通过，原因是已有 profile 配置，而非本仓库技能：

1. 任务启动时报 `deployment:persona-suffix` 中未知变量 `{{cwd}}`。
2. `--dump-config` 另报 patch entry `web-search-local` 不存在。

因此，本记录不把 DSH 提供器通过写成“DSH 全链路会话通过”。仓库技能本身已被 DSH 原生发现并加载；完整使用仍需另行修复上述既有 DSH profile。未擅自修改该全局环境。

## 复测方式

在仓库根目录运行：

```bash
python3 .agents/skills/hetong-craft/scripts/test_hetong_redline.py
python3 .agents/skills/hetong-craft/scripts/test_word_reader.py
python3 .agents/skills/hetong-craft/scripts/test_office_kit.py
python3 .agents/skills/hetong-craft/scripts/test_review_audit.py
python3 .agents/skills/hetong-review/scripts/test_review_kit.py
python3 .agents/skills/hetong-craft/scripts/office_kit.py doctor --full
```

Pi 在新会话中使用 `/skill:hetong-craft`；Codex、DSH 通过各自技能选择方式或自然语言指定技能名。目录、frontmatter 或路由变化后，应重新执行三个宿主的加载探针，不能只以文件存在作为通过依据。
