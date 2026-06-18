param(
    [string]$OutputPath = "release\EqNB.dotm"
)

$ErrorActionPreference = "Stop"

$root = (Resolve-Path "$PSScriptRoot\..").Path
$modulePath = Join-Path $root "vba\EquationNumbering.bas"
$ribbonPath = Join-Path $root "ribbon\customUI14.xml"
$resolvedOutput = [System.IO.Path]::GetFullPath((Join-Path $root $OutputPath))
$outputDirectory = Split-Path -Parent $resolvedOutput
New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null

function Set-ZipTextEntry {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName,
        [string]$Content
    )

    $existing = $Archive.GetEntry($EntryName)
    if ($null -ne $existing) {
        $existing.Delete()
    }

    $entry = $Archive.CreateEntry($EntryName)
    $stream = $entry.Open()
    $writer = New-Object System.IO.StreamWriter($stream, [System.Text.UTF8Encoding]::new($false))
    try {
        $writer.Write($Content)
    }
    finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Get-ZipTextEntry {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        throw "Missing package entry: $EntryName"
    }

    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
        $stream.Dispose()
    }
}

function Add-RibbonCustomUi {
    param(
        [string]$DotmPath,
        [string]$CustomUiPath
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $archive = [System.IO.Compression.ZipFile]::Open($DotmPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        Set-ZipTextEntry -Archive $archive -EntryName "customUI/customUI14.xml" -Content (Get-Content -Raw -Encoding UTF8 $CustomUiPath)

        [xml]$relationships = Get-ZipTextEntry -Archive $archive -EntryName "_rels/.rels"
        $relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
        $hasRibbonRelationship = $false
        foreach ($relationship in $relationships.Relationships.Relationship) {
            if ($relationship.Target -eq "customUI/customUI14.xml") {
                $hasRibbonRelationship = $true
            }
        }
        if (-not $hasRibbonRelationship) {
            $nextId = 1
            foreach ($relationship in $relationships.Relationships.Relationship) {
                if ($relationship.Id -match "^rId(\d+)$") {
                    $nextId = [Math]::Max($nextId, [int]$Matches[1] + 1)
                }
            }
            $node = $relationships.CreateElement("Relationship", $relationshipNamespace)
            $node.SetAttribute("Id", "rId$nextId")
            $node.SetAttribute("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
            $node.SetAttribute("Target", "customUI/customUI14.xml")
            $relationships.Relationships.AppendChild($node) | Out-Null
            Set-ZipTextEntry -Archive $archive -EntryName "_rels/.rels" -Content $relationships.OuterXml
        }
    }
    finally {
        $archive.Dispose()
    }
}

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
    $template.Close($false)
    $template = $null

    Add-RibbonCustomUi -DotmPath $resolvedOutput -CustomUiPath $ribbonPath
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
