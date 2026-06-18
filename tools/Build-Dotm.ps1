param(
    [string]$OutputPath = "release\WordEquationNumbering.dotm"
)

$ErrorActionPreference = "Stop"

$root = (Resolve-Path "$PSScriptRoot\..").Path
$modulePath = Join-Path $root "vba\EquationNumbering.bas"
$resolvedOutput = [System.IO.Path]::GetFullPath((Join-Path $root $OutputPath))
$outputDirectory = Split-Path -Parent $resolvedOutput
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

$word = $null
$template = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $template = $word.Documents.Add()

    try {
        $project = $template.VBProject
        if ($null -eq $project) {
            throw "VBProject is null"
        }
        $project.VBComponents.Import($modulePath) | Out-Null
    }
    catch {
        throw @"
Word blocked programmatic VBA project access.

To build the .dotm automatically, open Word and enable:
File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model

Original error: $($_.Exception.Message)
"@
    }

    # wdFormatXMLTemplateMacroEnabled = 15
    $template.SaveAs2($resolvedOutput, 15)
    Write-Host "Created $resolvedOutput"
}
finally {
    if ($template -ne $null) {
        $template.Close($false)
    }
    if ($word -ne $null) {
        $word.Quit()
    }
}
