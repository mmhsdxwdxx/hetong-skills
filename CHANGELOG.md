# Changelog

## 2.0.0-dev.2（develop，2026-10-31）

- 增加Windows PowerShell安装／只检查入口，按明确开关安装缺失Python和LibreOffice，隔离依赖及检查证据。
- 增加虚构Word修订、批注、XQY报告与真实PDF字体检查；doctor可保留样张和JSON报告。
- 支持非标准LibreOffice可执行路径；Windows渲染进程不继承Unix字体配置，快审CLI统一UTF-8输出。
- 纠正Windows系统字体与macOS/Linux fontconfig的边界，不假定Codex捆绑LibreOffice。
- CI增加Windows实际LibreOffice渲染及PowerShell 5.1端到端检查；不刷新已有线程，不发布稳定版。

## 2.0.0-dev.1（develop，2026-10-31）

- 将独立hetong-server撤并至hetong-craft；甲乙方分类条款、受理及未留痕检查统一维护。旧入口属于兼容性变更，升级时先备份旧安装。
- 新增hetong-fast-xqy-v1：按甲乙方立场生成详细Word审核报告及原稿批注副本，以R编号衔接。
- 固化履约时间轴规范、A4固定列宽、应付／已付／待退／资金占用口径及离线Markdown／HTML模板。
- 同步原生受理、按需协作和克制批注规范；提供Claude/Pi项目适配入口。
- 更新WorkBuddy构建、安装提示及CI到新架构。开发版只供新会话按项目路径加载，不自动刷新既有会话；未变更稳定版标签或创建Release。

## 1.3.0-dev.1

- 增加 WorkBuddy 原生安装包：`hetong.skill` 单包上传、三套单技能包与合并 zip；安装脚本不修改 WorkBuddy 原有资产。
- 增加 WorkBuddy 验证：官方 `quick_validate.py` 通过，三套技能同级安装后脚本入口可用。


- 增加 Windows 11 原生适配：`install.ps1` / `install.cmd`、`Scripts\python.exe` 虚拟环境路径、LibreOffice Windows 路径、`msvcrt` 文件锁和系统字体支持。
- 增加 Windows CI 与平台路径回归；macOS Office 渲染回归保持通过。
- 未改变 `main` / `v1.2.0` 稳定版。

## 1.2.0

- 为 `hetong-server` 增加两版近似 DOCX 的未留痕改动检测：直接改文、留痕改文、历史/批注减少、隐藏文字和媒体变化。
- 新增 `hidden_change.py`、检测报告模板和 8 项虚构回归测试；工具只读，不自动接受或拒绝修订。
- 保留甲方 A01—A04 与乙方 B01—B11 分列；公开版本继续移除真实交易参数、源文档指纹和具体署名。

## 1.1.0

- 为 `hetong-server` 增加甲方视角 A01—A04：保函受益人/生效、BIS 与政府禁令、保函有效期与提取、资金共管账户及 U 盾权限。
- 乙方 B01—B11 与甲方 A01—A04 分列；按我方立场读取，不叠加适用。
- 公开发布版继续移除真实交易排期、历史金额/台数、源文档指纹和具体署名。

## 1.0.0

- 发布 `hetong-craft`、`hetong-review`、`hetong-server` 三套技能。
- 包含 DOCX 修订/批注、Word 深读与时间线、Office 无头预览、审核工具与虚构回归测试。
- 公开版本已脱敏：移除真实交易排期、历史金额/台数、源文档指纹和具体署名。
