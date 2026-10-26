#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Target
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Source = Join-Path $Root ".agents\skills"

if (-not (Test-Path -LiteralPath $Source -PathType Container)) {
    throw "Skill source not found: $Source"
}
if (Test-Path -LiteralPath $Target -PathType Leaf) {
    throw "Target must be a project directory, not a file: $Target"
}
if (-not (Test-Path -LiteralPath $Target -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $Target | Out-Null
}

$Destination = Join-Path $Target ".agents\skills"
New-Item -ItemType Directory -Force -Path $Destination | Out-Null

# Robocopy is part of Windows. Exit codes below 8 mean success/copy completed.
& robocopy $Source $Destination /E /COPY:DAT /R:2 /W:1 /NFL /NDL /NJH /NJS /NP
if ($LASTEXITCODE -ge 8) {
    throw "Robocopy failed with exit code $LASTEXITCODE"
}

Write-Host "Installed skills to: $Destination"
Write-Host "Start Pi/Codex/DSH from the target project root and reload skills."
Write-Host "This installer overwrites same-named files but does not delete extra files."
exit 0
