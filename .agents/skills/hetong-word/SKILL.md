---
name: hetong-word
description: 处理合同 Word 文件的读取、修订历史深读、真实留痕修改、范围批注、成品验证和可选页面预览。用于 hetong 系列需要操作 DOCX 或制作 Word 交付物时；合同判断、条款选择和商业条件仍由上游技能负责。
---

# 合同 Word 操作

**v2.0.2审阅规则**：先执行[统一审阅输出规则](references/review-rules.md)。批注仅针对真实问题，通常一两句，不转介文件；双交付批注与详细意见使用同号。一般任务中的明确错误直接在副本批注并最小修订，中英对照以中文为准并同步英文；简要报告总篇幅不超过两页。本次明确只读/只批注限制及未确认商业条件仍须遵守。


`hetong-craft`、`hetong-review` 和 `hetong-fast-xqy-v1` 共用的 Word 操作子技能。
实际需要读取、生成、批注、修订或验证 Word 时加载，并向用户说明“本次使用 `hetong-word` 处理 Word 文件”。纯文字讨论无需加载。调用本技能不等于启动子代理；默认由当前主代理执行。

## 初始化：检测在先，组件按需

首次使用、切换运行环境、依赖变化或故障时，按[初始化说明](references/environment-init.md)检查本次能力。已有检查结果且环境未变时直接复用，不每次安装或提问。

```bash
python3 <本技能>/scripts/init.py --profile core
```

- `core`：提取、深读、真实修订、批注及结构验证，仅需 Python 3.10+ 标准库。
- `generate`：制作 Word 报告等，按需使用 `python-docx`。
- `render`：页面预览，需要 PDF/图像依赖及 LibreOffice。
- `office`：兼容原 Office 完整工具链，仅跨格式精修或完整环境验收需要。

无安装参数的初始化只检测，不安装包、不创建虚拟环境、不改全局配置。先复用宿主已提供或已有的合适 Python；只有当前任务需要且已有授权时才用 `--install-python-deps`、`--install-libreoffice` 分别安装。技能安装与软件安装分开；LibreOffice 缺失不阻断 `core` / `generate`。

## 内容与操作边界

上游负责合同立场、事实、法律判断、条款及批注准确文字、已确认商业条件和允许操作。直接调用本技能也须有明确的操作范围；缺少的商业条件不得自行填写。

- 源件只读、默认另存；保留既有修订、批注和样式，禁止为绕过工具限制接受/拒绝全部历史。
- 沿用本次有效授权，不重复索要批准。一般审核可按统一规则纠正明确错误；本次明确只读或仅批注时不允许改正文。真实合同的受理和类别选择仍使用 [craft 受理流程](../hetong-craft/references/host-intake.md)，XQY沿用其专用授权，不另造回执。
- 执行[审阅隐私](references/review-privacy.md)：批注、回复、修订文字及理由独立完整，不含内部报告名称、路径、链接或“详见/参见”。双交付允许独立完整意见前显示同一R编号；技术映射单独记录。
- 新增批注仍由上游按[触发标准](../hetong-craft/references/comment-policy.md)判断。操作技能不增加新的审核意见或改写已定措辞。
- 默认用 WPS 打开和原生验收，只有用户明确指定 Microsoft Word 程序时才使用它；见[客户端与显示](references/native-review-display.md)。不修改系统文件关联。

## 按任务加载

| 操作 | 流程与工具 |
|---|---|
| 提取、定位、普通清洁稿、新建 DOCX | [DOCX 工作流](references/docx-workflow.md)，`scripts/docx_tool.py` |
| 真实修订、范围批注 | [红线与批注](references/docx-redline.md)，`scripts/hetong_redline.py` |
| 已有修订/批注/回复的文件 | [Word 深读](references/word-deep-read.md)，`scripts/word_reader.py` |
| 批注选区与文字质量 | [批注工作流](references/comment-workflow.md) |
| 历史重叠、复杂结构、清洁稿派生 | [工具边界与替代路径](references/word-operation-paths.md) |
| 两版未留痕改动 | [未留痕检查](references/hidden-change-review.md)，`scripts/hidden_change.py` |
| 成品历史、锚点与包结构核对 | `scripts/audit_docx_review.py`；用 `--before` 指定基线 |
| 简要报告两页检查 | `scripts/verify_brief_report.py --source 报告.docx --manifest 渲染清单.json`，核对实际PDF页数及指纹 |
| 可选预览及共用 Office 工具 | [预览与 Office 工作流](references/office-workflow.md)，`scripts/office_kit.py` |

仅加载本次需要的参考。含历史审阅的文件在编辑前核对原文、接受视图、标记视图、时间线和批注线程；这些分析视图不写回原件，作者活动时段不等于谈判轮次。

## 执行与验收

定位与核对基线 → 生成计划/`--dry-run` → 另存 → `verify` → 重新提取并核对授权目标。
红线验证分别核对仅接受和仅拒绝本轮修订的结果；批注稿验证正文及历史不变。原件指纹、历史保留、批注隐私和锚点为必检项。

有版面风险或交付规格要求时做页面视觉复核。LibreOffice 可选安装，不表示视觉复核可省略；可用 WPS 等本地路径验证。无可用路径时明确说明未验证部分，不宣称完整验收。结构检查、渲染、WPS/Word 原生另存往返分别记录。

内容判断和最终验收由当前主代理负责。确有必要且获准委派时才使用 [Word 操作交接](../hetong-craft/references/word-operation-agent.md)，不因加载本技能自动委派。
