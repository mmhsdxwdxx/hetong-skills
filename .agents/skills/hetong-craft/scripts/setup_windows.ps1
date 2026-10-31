#requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$InstallPrerequisites,
    [switch]$CheckOnly,
    [string]$PythonExecutable,
    [string]$LibreOfficeExecutable,
    [string]$Venv,
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Find-Python312 {
    param([string]$Explicit)
    $candidates = @()
    if ($Explicit) { $candidates = @($Explicit) }
    else {
        $launcher = Get-Command py.exe -ErrorAction SilentlyContinue
        if ($launcher) {
            try {
                $resolved = & $launcher.Source -3.12 -c 'import sys; print(sys.executable)' 2>$null
                if ($LASTEXITCODE -eq 0) { $candidates += ($resolved | Select-Object -Last 1) }
            } catch { }
        }
        foreach ($base in @($env:LOCALAPPDATA, $env:ProgramFiles)) {
            if ($base) {
                $candidates += (Join-Path $base 'Programs\Python\Python312\python.exe')
                $candidates += (Join-Path $base 'Python312\python.exe')
            }
        }
        foreach ($name in @('python.exe', 'python3.exe')) {
            $command = Get-Command $name -ErrorAction SilentlyContinue
            if ($command) { $candidates += $command.Source }
        }
    }
    foreach ($candidate in $candidates) {
        if (-not $candidate -or -not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
        try {
            $probe = & $candidate -c "import sys,struct; print('ok' if sys.version_info[:2]==(3,12) and struct.calcsize('P')==8 else 'unsupported')" 2>$null
            if ($LASTEXITCODE -eq 0 -and ($probe | Select-Object -Last 1) -eq 'ok') {
                return (Resolve-Path -LiteralPath $candidate).Path
            }
        } catch { }
    }
    if ($Explicit) { throw '指定Python不可用；需要64位Python 3.12。' }
    return $null
}

function Find-LibreOffice {
    param([string]$Explicit)
    if ($Explicit) {
        if (-not (Test-Path -LiteralPath $Explicit -PathType Leaf)) { throw '指定LibreOffice路径不存在。' }
        return (Resolve-Path -LiteralPath $Explicit).Path
    }
    if ($env:CONTACT_LIBREOFFICE) { return Find-LibreOffice -Explicit $env:CONTACT_LIBREOFFICE }
    $command = Get-Command soffice.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:LOCALAPPDATA)) {
        if ($base) {
            $candidate = Join-Path $base 'LibreOffice\program\soffice.exe'
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
        }
    }
    return $null
}

function Install-RequiredPackage {
    param([string]$Id)
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) { throw '没有winget。请从官方安装Python 3.12及LibreOffice，随后重跑；不自动安装其他包管理器。' }
    & $winget.Source install --id $Id --exact --source winget --accept-package-agreements --accept-source-agreements --disable-interactivity
    if ($LASTEXITCODE -ne 0) { throw "winget安装失败：$Id；退出码$LASTEXITCODE。请完成安装后重跑。" }
}

function Invoke-Setup {
    if ($env:OS -ne 'Windows_NT') { throw '此脚本仅用于Windows；其他平台使用office_kit.py setup/doctor。' }
    if ($CheckOnly -and $InstallPrerequisites) { throw 'CheckOnly不能与InstallPrerequisites同时使用。' }
    if (-not $Venv) { $Venv = Join-Path $env:LOCALAPPDATA 'contact-office\venv' }
    if (-not $OutputDirectory) {
        $stamp = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
        $OutputDirectory = Join-Path (Get-Location).Path ("产出\环境检查-" + $stamp)
    }
    if (Test-Path -LiteralPath $OutputDirectory) { throw '检查输出目录已存在；请使用新目录以保留历史结果。' }
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
    $saved = @{}
    foreach ($name in @('CONTACT_OFFICE_VENV','CONTACT_LIBREOFFICE','PYTHONUTF8','PYTHONIOENCODING')) {
        $saved[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
    }
    $result = [ordered]@{ ok=$false; check_only=[bool]$CheckOnly; setup_performed=$false; artifacts=$OutputDirectory; native_word_wps_roundtrip='not performed' }
    try {
        $python = Find-Python312 -Explicit $PythonExecutable
        if (-not $python -and $InstallPrerequisites) {
            Install-RequiredPackage -Id 'Python.Python.3.12'
            $python = Find-Python312
        }
        if (-not $python) { throw '未找到64位Python 3.12。安装后重跑，或传-PythonExecutable完整路径。' }
        $lo = Find-LibreOffice -Explicit $LibreOfficeExecutable
        if (-not $lo -and $InstallPrerequisites) {
            Install-RequiredPackage -Id 'TheDocumentFoundation.LibreOffice'
            $lo = Find-LibreOffice
        }
        if (-not $lo) { throw '未找到LibreOffice。安装Windows版后重跑，或传-LibreOfficeExecutable完整路径。' }
        $env:CONTACT_OFFICE_VENV = $Venv
        $env:CONTACT_LIBREOFFICE = $lo
        $env:PYTHONUTF8 = '1'
        $env:PYTHONIOENCODING = 'utf-8'
        $kit = Join-Path $PSScriptRoot 'office_kit.py'
        if (-not $CheckOnly) {
            & $python $kit setup --venv $Venv
            if ($LASTEXITCODE -ne 0) { throw 'Office隔离依赖配置失败；查看上方错误。' }
            $result.setup_performed = $true
        }
        $officePython = Join-Path $Venv 'Scripts\python.exe'
        if (-not (Test-Path -LiteralPath $officePython -PathType Leaf)) { throw '隔离环境尚未建立；先不带-CheckOnly执行一次。' }
        $result['python'] = $officePython
        $result['libreoffice'] = $lo
        & $officePython $kit doctor --full --artifacts (Join-Path $OutputDirectory '基础样张') --out (Join-Path $OutputDirectory 'doctor.json')
        if ($LASTEXITCODE -ne 0) { throw '基础渲染检查失败；查看doctor.json。' }
        & $officePython (Join-Path $PSScriptRoot 'office_smoke.py') --out (Join-Path $OutputDirectory 'Word样张') --require-fonts
        if ($LASTEXITCODE -ne 0) { throw 'Word修订/批注/快审或宋体黑体检查失败；查看Word样张\smoke-result.json。' }
        $result.ok = $true
        Write-Host "检查通过。请用Word/WPS打开保留样张，人工核对批注、修订及分页：$OutputDirectory"
        Write-Host '未刷新已有线程；新开项目会话使用新版。'
    } catch {
        $result['error'] = $_.Exception.Message
        throw
    } finally {
        $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'setup-result.json') -Encoding UTF8
        foreach ($name in $saved.Keys) { [Environment]::SetEnvironmentVariable($name, $saved[$name], 'Process') }
    }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-Setup }
