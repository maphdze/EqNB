param(
    [string]$SourcePath = "release\EqNB.dotm",
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$root = (Resolve-Path "$PSScriptRoot\..").Path
$resolvedSource = [System.IO.Path]::GetFullPath((Join-Path $root $SourcePath))

if (-not (Test-Path -LiteralPath $resolvedSource)) {
    throw "EqNB.dotm was not found at: $resolvedSource. Run tools\Build-Dotm.ps1 first."
}

$wordProcesses = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($wordProcesses -and -not $Force) {
    throw "Word is running. Close all Word windows first, or rerun with -Force."
}

$startupPath = Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
New-Item -ItemType Directory -Force -Path $startupPath | Out-Null

$targetPath = Join-Path $startupPath "EqNB.dotm"
Copy-Item -LiteralPath $resolvedSource -Destination $targetPath -Force

Write-Host "Installed EqNB.dotm to:"
Write-Host $targetPath
Write-Host ""
Write-Host "Restart Word, then open the EqNB tab."
