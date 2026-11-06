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
$invokeNow = $MyInvocation.InvocationName -ne '.'
. (Join-Path $PSScriptRoot '../../hetong-word/scripts/setup_windows.ps1') @PSBoundParameters
if ($invokeNow) { Invoke-Setup }
