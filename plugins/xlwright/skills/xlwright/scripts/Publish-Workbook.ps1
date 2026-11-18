<#
.SYNOPSIS
Updates a work book from its VBA (and optional ribbon) sources and publishes a VBA-locked copy.

.DESCRIPTION
The sources are read from the work book's folder: vba/ (required) and ribbon/customUI14.xml
(embedded only when present). The default output file name is the work book's name, or the folder
name + .xlsm when the work book is named workbook.xlsm.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OutDir,
    [ValidateNotNullOrEmpty()][string]$FileName,
    [switch]$Force
)
$ErrorActionPreference = 'Stop'

$workbook = Get-Item -LiteralPath $Path
if ($workbook -isnot [IO.FileInfo] -or $workbook.Extension -ne '.xlsm') {
    throw "Expected an .xlsm file: $Path"
}

$sourceDir = Join-Path $workbook.DirectoryName 'vba'
$ribbonPath = Join-Path $workbook.DirectoryName 'ribbon/customUI14.xml'
if (!(Test-Path -LiteralPath $sourceDir -PathType Container)) { throw "Missing VBA directory: $sourceDir" }
$hasRibbon = Test-Path -LiteralPath $ribbonPath -PathType Leaf

if (!$PSBoundParameters.ContainsKey('FileName')) {
    $FileName = if ($workbook.Name -eq 'workbook.xlsm') { $workbook.Directory.Name + '.xlsm' } else { $workbook.Name }
}
if ([IO.Path]::GetExtension($FileName) -ne '.xlsm' -or
    $FileName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) {
    throw "Expected a file name ending in .xlsm: $FileName"
}
$provider = $null
$drive = $null
$outputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
    $OutDir, [ref]$provider, [ref]$drive)
if ($provider.Name -ne 'FileSystem') { throw "Expected a filesystem directory: $OutDir" }
$outputPath = [IO.Path]::Combine($outputDirectory, $FileName)
if ([string]::Equals($outputPath, $workbook.FullName, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Output must be separate from the work book: $outputPath"
}
if ([IO.Directory]::Exists($outputPath)) { throw "Output is a directory: $outputPath" }
if (!$Force -and (Test-Path -LiteralPath $outputPath)) {
    throw "Output already exists: $outputPath. Use -Force to overwrite it."
}

$publishId = '.excel-publish-' + [guid]::NewGuid().ToString('N')
$updatedWorkbook = Join-Path $workbook.DirectoryName ($publishId + '.xlsm')
$pendingOutput = [IO.Path]::Combine($outputDirectory, $publishId + '.xlsm')
try {
    # Keep the work book intact until the imports and the release lock succeed.
    [IO.File]::Copy($workbook.FullName, $updatedWorkbook)
    & (Join-Path $PSScriptRoot 'Import-Vba.ps1') -Path $updatedWorkbook -SourceDir $sourceDir | Write-Verbose
    if ($hasRibbon) {
        & (Join-Path $PSScriptRoot 'Add-CustomUI.ps1') -Path $updatedWorkbook -XmlPath $ribbonPath | Write-Verbose
    }

    [void][IO.Directory]::CreateDirectory($outputDirectory)
    [IO.File]::Copy($updatedWorkbook, $pendingOutput)
    & (Join-Path $PSScriptRoot 'Lock-VbaProject.ps1') -Path $pendingOutput
    [IO.File]::Replace($updatedWorkbook, $workbook.FullName, [NullString]::Value)

    # Publish within the output directory so a partial copy is never visible.
    if ($Force -and [IO.File]::Exists($outputPath)) {
        [IO.File]::Replace($pendingOutput, $outputPath, [NullString]::Value)
    }
    else {
        [IO.File]::Move($pendingOutput, $outputPath)
    }
    $outputPath
}
finally {
    Remove-Item -LiteralPath $updatedWorkbook, $pendingOutput -Force -ErrorAction SilentlyContinue
}
