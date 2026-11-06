# Windows 初始化与按需组件

技能安装只提供技能文件，不自动安装 Python、Python 包、LibreOffice 或字体。首次实际操作 Word 时按[初始化流程](environment-init.md)检测所需能力。

```powershell
# 默认 core：只检测，无 LibreOffice 要求
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .agents\skills\hetong-word\scripts\setup_windows.ps1 -CheckOnly
# 已获准配置 Word 生成依赖
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .agents\skills\hetong-word\scripts\setup_windows.ps1 -Profile generate -InstallPythonDependencies
# 需要页面预览且已有安装授权：分别选择两类组件
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .agents\skills\hetong-word\scripts\setup_windows.ps1 -Profile render -InstallPythonDependencies -InstallLibreOffice
# 仅完整 Office 验收使用；需已具备依赖
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .agents\skills\hetong-word\scripts\setup_windows.ps1 -Profile office -FullValidation -OutputDirectory 产出\新环境检查
```

Python 3.10+ 即可执行基础操作；可用 `-PythonExecutable` 指定。只有显式 `-InstallPython` 才通过 winget 安装 Python 3.12。`-Venv` 选择隔离环境，`-LibreOfficeExecutable` 指定已有安装；显式路径无效直接报错。

旧 craft 路径继续转发。旧 `-InstallPrerequisites` 明确选择完整 Office 安装及验收，可用 `-OutputDirectory` 指定新目录，未指定则自动创建带时间戳的检查目录；新任务优先使用细分参数。`-CheckOnly` 与所有安装开关互斥，不调用 pip 或 winget。只检测默认不写文件；指定 `-OutputDirectory` 才保存 JSON，不覆盖旧证据。

ExecutionPolicy 仅影响本次进程。脚本不改系统文件关联、不重启或刷新已有聊天、不自动补装字体。中文字体从 Windows 官方可选功能或其他授权来源配置；安装失败保留错误，文件已生成不代表验收通过。原生 WPS 验收与 XML/PDF 检查分开记录。
