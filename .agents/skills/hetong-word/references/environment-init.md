# hetong-word 初始化与可选组件

初始化是技能工作流的一步，不是宿主安装技能时会自动执行的钩子。安装技能只提供说明和脚本；代理首次用到能力时运行检测，并按用户已授权的范围配置缺失组件。不存在“读取 SKILL.md 即允许安装软件”的隐含授权。

## 能力分层

| profile | 用途 | 必需组件 |
|---|---|---|
| `core`（默认） | DOCX 提取、深读、修订、批注、结构审计 | Python 3.10+；无第三方包、无 LibreOffice |
| `generate` | Word 报告生成 | `python-docx` 及其锁定依赖 |
| `render` | Word 转 PDF、页面图片与视觉预览 | generate + PDF/图像包 + LibreOffice |
| `office` | 兼容 PPTX/XLSX 精修及原 `doctor --full` | 原完整 Office 锁定依赖 + LibreOffice |

只按本次任务所需 profile 检测。初始化输出 JSON，含实际 Python 路径、缺失组件、所选 profile 是否可用、安装动作和未执行的原生验收。`ok` 只代表该 profile 的依赖检查通过，不代表文档内容或客户端验收通过。

## 常用命令

```bash
# 默认只检查，不安装、不写配置
python3 <hetong-word>/scripts/init.py
python3 <hetong-word>/scripts/init.py --profile generate
python3 <hetong-word>/scripts/init.py --profile render

# 复用宿主提供的 Python；显式路径无效时报错，不换成其他环境
python3 <hetong-word>/scripts/init.py --profile generate --python /absolute/path/python

# 已获准配置 Python 依赖时：仅向隔离环境安装所选能力的锁定包
python3 <hetong-word>/scripts/init.py --profile generate --install-python-deps

# 已获准安装 LibreOffice 时单独选择；可和 Python 依赖开关组合
python3 <hetong-word>/scripts/init.py --profile render --install-libreoffice

# 自定义隔离环境与已有 LibreOffice，不修改系统默认配置
python3 <hetong-word>/scripts/init.py --profile render --venv /absolute/venv --libreoffice /absolute/soffice
```

`--check-only` 显式表示只检查，不能与任一安装开关组合。无安装开关时行为相同。先检查当前 Python、`CONTACT_OFFICE_VENV` 和已有用户环境；显式 `--python` 或 `--venv` 只检查指定对象，不回退。返回的 `python` 路径用于后续生成和渲染命令。

默认隔离环境沿用用户数据目录：Windows `%LOCALAPPDATA%\contact-office\venv`；macOS/Linux `${XDG_DATA_HOME:-~/.local/share}/contact-office/venv`。环境、缓存和预览均不得放入技能包。已有可用依赖直接复用；仅缺失时安装所选组件，不主动升级其他包。

## 平台适配和验证

- Windows：可用 [PowerShell 初始化入口](../scripts/setup_windows.ps1)，默认 `core`，`-Profile` 选择能力，`-InstallPythonDependencies` / `-InstallLibreOffice` 分别授权安装。Python 缺失时只有 `-InstallPython` 才调用 winget。旧 `-InstallPrerequisites` 保留为显式请求完整 Office 环境的兼容参数。
- macOS：复用已有 Python；选择安装 LibreOffice 时使用已有 Homebrew。不自动安装 Homebrew。
- Linux：复用已有 Python；只有已选安装且具备 root 权限时使用已有 apt-get/dnf。其他情况给出组件缺口，由用户或管理员安装，不自行 sudo。
- 中文字体和 WPS/Word 客户端单独核验，不自动下载字体或修改默认打开程序。系统安装可能需要交互或管理员权限；失败保留原错误，不视为完成。

环境变化后，对实际启用能力做回归：core 跑红线、深读和审阅审计；generate 生成并回读虚构报告；render 实际渲染虚构 DOCX 并检查中文文本层/页面图。选择完整 office 后才运行 `office_kit.py doctor --full` 和跨格式回归。缺少 LibreOffice 不导致 core 检查失败，也不能伪称已完成页面或原生往返验收。
