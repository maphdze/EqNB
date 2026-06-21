param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$wordProcesses = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($wordProcesses -and -not $Force) {
    throw "Word is running. Close all Word windows first, or rerun with -Force."
}

$startupPath = Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
$targetPath = Join-Path $startupPath "EqNB.dotm"

if (Test-Path -LiteralPath $targetPath) {
    Remove-Item -LiteralPath $targetPath -Force
    Write-Host "Removed EqNB.dotm from:"
    Write-Host $targetPath
}
else {
    Write-Host "EqNB.dotm was not found in:"
    Write-Host $startupPath
}

Write-Host ""
Write-Host "Restart Word to unload EqNB."
