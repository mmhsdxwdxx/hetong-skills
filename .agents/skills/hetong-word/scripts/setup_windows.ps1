#requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('core','generate','render','office')][string]$Profile = 'core',
    [switch]$InstallPython,
    [switch]$InstallPythonDependencies,
    [switch]$InstallLibreOffice,
    [switch]$InstallPrerequisites,
    [switch]$CheckOnly,
    [switch]$FullValidation,
    [string]$PythonExecutable,
    [string]$LibreOfficeExecutable,
    [string]$Venv,
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$script:HetongWordScriptRoot = $PSScriptRoot

function Find-Python {
    param([string]$Explicit)
    $candidates = @()
    if ($Explicit) { $candidates = @($Explicit) }
    else {
        $launcher = Get-Command py.exe -ErrorAction SilentlyContinue
        if ($launcher) {
            try {
                $resolved = & $launcher.Source -3 -c 'import sys; print(sys.executable)' 2>$null
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
            $probe = & $candidate -c "import sys,struct; print('ok' if sys.version_info >= (3,10) else 'unsupported')" 2>$null
            if ($LASTEXITCODE -eq 0 -and ($probe | Select-Object -Last 1) -eq 'ok') {
                return (Resolve-Path -LiteralPath $candidate).Path
            }
        } catch { }
    }
    if ($Explicit) { throw '指定Python不可用；需要Python 3.10+。' }
    return $null
}

function Invoke-Setup {
    if ($env:OS -ne 'Windows_NT') { throw '此入口仅用于Windows；其他平台使用hetong-word/scripts/init.py。' }
    if ($CheckOnly -and ($InstallPython -or $InstallPythonDependencies -or $InstallLibreOffice -or $InstallPrerequisites)) {
        throw 'CheckOnly不能与安装开关组合。'
    }
    if ($InstallPrerequisites) {
        $Profile = 'office'
        $InstallPython = $true
        $InstallPythonDependencies = $true
        $InstallLibreOffice = $true
        $FullValidation = $true
    }
    if ($FullValidation -and $Profile -ne 'office') { throw 'FullValidation仅适用于显式选择office能力。' }
    if ($FullValidation -and -not $OutputDirectory) {
        $stamp = (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
        $OutputDirectory = Join-Path (Get-Location).Path ('产出\环境检查-' + $stamp)
    }
    if ($OutputDirectory -and (Test-Path -LiteralPath $OutputDirectory)) { throw '输出目录已存在；不覆盖旧检查证据。' }
    $python = Find-Python -Explicit $PythonExecutable
    if (-not $python -and $InstallPython) {
        $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
        if (-not $winget) { throw '未找到winget；请从官方来源安装Python 3.10+后重跑。' }
        & $winget.Source install --id Python.Python.3.12 --exact --source winget --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -ne 0) { throw 'Python安装失败。' }
        $python = Find-Python
    }
    if (-not $python) { throw '未找到Python 3.10+；可指定-PythonExecutable或在获准后使用-InstallPython。' }
    $arguments = @((Join-Path $script:HetongWordScriptRoot 'init.py'),'--profile',$Profile)
    if ($Venv) { $arguments += @('--venv',$Venv) }
    elseif ($PythonExecutable -and -not $InstallPythonDependencies) { $arguments += @('--python',$python) }
    if ($LibreOfficeExecutable) { $arguments += @('--libreoffice',$LibreOfficeExecutable) }
    if ($CheckOnly) { $arguments += '--check-only' }
    if ($InstallPythonDependencies) { $arguments += '--install-python-deps' }
    if ($InstallLibreOffice) { $arguments += '--install-libreoffice' }
    $raw = & $python @arguments
    $initExit = $LASTEXITCODE
    $result = ($raw -join "`n") | ConvertFrom-Json
    $result | Add-Member -NotePropertyName 'capability_ok' -NotePropertyValue ([bool]$result.ok)
    $result | Add-Member -NotePropertyName 'full_validation' -NotePropertyValue 'not_requested'
    $validationError = $null
    if ($FullValidation -and $initExit -eq 0) {
        $savedLo = $env:CONTACT_LIBREOFFICE
        try {
            $env:CONTACT_LIBREOFFICE = $result.libreoffice.path
            & $result.python (Join-Path $script:HetongWordScriptRoot 'office_kit.py') doctor --full --artifacts (Join-Path $OutputDirectory '基础样张') --out (Join-Path $OutputDirectory 'doctor.json') | Out-Host
            if ($LASTEXITCODE -ne 0) { throw '完整Office渲染检查失败。' }
            & $result.python (Join-Path $script:HetongWordScriptRoot 'office_smoke.py') --out (Join-Path $OutputDirectory 'Word样张') --require-fonts | Out-Host
            if ($LASTEXITCODE -ne 0) { throw 'Word样张或字体检查失败。' }
            $result.full_validation = 'passed'
        } catch {
            $result.ok = $false
            $result.full_validation = 'failed'
            $validationError = $_.Exception.Message
            $result | Add-Member -NotePropertyName 'validation_error' -NotePropertyValue $validationError
        } finally { $env:CONTACT_LIBREOFFICE = $savedLo }
    }
    $json = $result | ConvertTo-Json -Depth 8
    if ($OutputDirectory) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        $json | Set-Content -LiteralPath (Join-Path $OutputDirectory 'setup-result.json') -Encoding UTF8
    }
    $json | Write-Output
    if ($initExit -ne 0) { throw '所选能力初始化失败；查看JSON中的missing或error。' }
    if ($validationError) { throw $validationError }
    if ($FullValidation) { Write-Host '请用WPS核对样张；原生客户端往返尚未执行。' }

}
if ($MyInvocation.InvocationName -ne '.') { Invoke-Setup }
