# Windows 11安装、升级与Office检查

适用于本地Windows 11和PowerShell 5.1／7。Word正文由Python处理；LibreOffice负责转换和预览，Word/WPS客户端负责最终原生显示验证。无需WSL，也不要执行macOS的brew、textutil或fontconfig修复命令。

## 首次安装：发给使用者

1. 准备Git、Node.js 22.20+和Codex。若缺Git或Node，可从其官方安装器安装，或在PowerShell分别执行下面两条命令；完成后重新打开终端，让PATH生效。

```powershell
winget install --id Git.Git --exact --source winget
winget install --id OpenJS.NodeJS.LTS --exact --source winget
```

2. 在“合同”仓库根目录打开PowerShell，安装develop分支的三个技能，然后运行Office安装检查。Windows使用npx.cmd，避免系统把npx解析为受执行策略限制的npx.ps1：

```powershell
npx.cmd --yes skills@latest add https://github.com/mmhsdxwdxx/hetong-skills/tree/develop --skill '*' --agent codex --yes
if ($LASTEXITCODE -ne 0) { throw '技能安装失败，请先检查网络和Git。' }
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .agents\skills\hetong-craft\scripts\setup_windows.ps1 -InstallPrerequisites
```

`-InstallPrerequisites`明确允许脚本在缺少时通过winget安装64位Python 3.12及Windows版LibreOffice，安装器可能弹出系统权限请求。已有程序不强制升级。没有winget时，脚本提示手动安装，不改用不明下载源。ExecutionPolicy只作用于本次PowerShell进程，不永久修改系统策略。

脚本安装或更新锁定的Python依赖到`%LOCALAPPDATA%\contact-office\venv`；运行两级检查，保留JSON、虚构DOCX、PDF和页面图到当前项目`产出\环境检查-时间戳-随机标识\`。失败返回非零状态，不把只生成文件当成功。它不重启Codex、Word或LibreOffice界面，不通知或刷新已有线程。

3. 成功后新开“合同”项目会话。用Word/WPS打开`Word样张`里的修订批注稿和快审报告，检查所有标记、批注选区、中文字体和分页；在临时副本中接受／拒绝修订并另存。该步骤是客户端验证，不能由PDF或CI代替。

## 日后升级

在同一个项目根目录重复安装命令即可更新develop中的技能。然后不带安装系统软件开关运行检查入口，它会更新锁定的Office依赖并重新检查：

```powershell
npx.cmd --yes skills@latest add https://github.com/mmhsdxwdxx/hetong-skills/tree/develop --skill '*' --agent codex --yes
if ($LASTEXITCODE -ne 0) { throw '技能升级失败。' }
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .agents\skills\hetong-craft\scripts\setup_windows.ps1
```

同名技能文件会被新版覆盖，不把个人修改写进技能安装目录；合同原件和成果放独立任务目录。旧会话不主动刷新；文件更新并不为主动重读磁盘的旧会话创建版本隔离。旧hetong-server目录如仍存在，应先备份并单独处理，避免与新架构同时加载。

## 只检查，不装依赖

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .agents\skills\hetong-craft\scripts\setup_windows.ps1 -CheckOnly
```

此模式不调用winget和pip，仍生成新的虚构检查证据，不处理真实合同。

## 自定义路径与故障定位

- 非标准安装位置：传`-PythonExecutable 'D:\Python312\python.exe'`、`-LibreOfficeExecutable 'D:\LibreOffice\program\soffice.exe'`。Python必须为64位3.12；无效显式路径直接报错，不悄悄选另一份安装。
- 自定义隔离环境：传`-Venv 'D:\hetong-runtime\venv'`。后续同样指定此路径；脚本不永久修改环境变量。
- 单独运行Python工具时可用当前进程的`CONTACT_OFFICE_VENV`和`CONTACT_LIBREOFFICE`指定路径。后者须指向可执行文件，不是目录。快审生成器使用隔离环境的`Scripts\python.exe`；不要仅因系统Python能启动就认定已装python-docx。
- 路径包含中文或空格时使用引号；脚本使用参数数组传递，不拼接shell命令。
- Windows采用系统字体；不要设置FONTCONFIG_FILE／FONTCONFIG_PATH。检查脚本要求样张实际使用宋体、黑体；缺失时从Windows“可选功能”的简体中文补充字体或有授权的字体来源补齐后重跑，不把Noto等替代字体称为宋体。
- `doctor.json`定位Python、LibreOffice、PDF及中文文本层；`Word样张\smoke-result.json`定位修订、批注、快审和PDF字体。`setup-result.json`是本次整体结果；`native_word_wps_roundtrip`仍为未执行。
- 关闭占用待处理文件的Word/WPS，或者处理独立副本。禁止为了“通过”删除历史修订、移除文档保护或降低检查。

参考：[winget安装参数](https://learn.microsoft.com/en-us/windows/package-manager/winget/install)、[LibreOffice下载](https://www.libreoffice.org/download/)、[skills CLI](https://github.com/vercel-labs/skills/blob/main/docs/commands.md)。具体包及运行环境以本次实际检查结果为准。
