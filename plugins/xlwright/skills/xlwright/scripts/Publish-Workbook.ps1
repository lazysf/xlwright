<#
.SYNOPSIS
Updates a work book from its VBA (and optional ribbon) sources and publishes a VBA-locked copy.

.DESCRIPTION
The sources are read from the work book's folder: vba/ (required) and ribbon/customUI14.xml
(embedded only when present). The work book is an .xlsm or an .xlsb. The default output file name is
the work book's name, or the folder name + the work book's extension when the work book is named
workbook.xlsm or workbook.xlsb. The extension of the output file name picks the format: .xlsm
(macro-enabled workbook), .xlsb (binary workbook) or .xlam (add-in). The work book keeps its format.

The output is saved with Workbook.RemovePersonalInformation set, which clears the author, the last author, the
manager, the company, the comment authors and the saved folder path. The work book keeps them. The output is read
back and rejected if one of the four properties remains or the flag is not set. -KeepPersonalInformation keeps
them and names the kept items in a warning.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OutDir,
    [ValidateNotNullOrEmpty()][string]$FileName,
    [switch]$KeepPersonalInformation,
    [switch]$Force
)
$ErrorActionPreference = 'Stop'

# The file formats of SaveAs (XlFileFormat) by the extensions an output can have.
$outputFormats = @{ '.xlsm' = 52; '.xlsb' = 50; '.xlam' = 55 }  # xlOpenXMLWorkbookMacroEnabled, xlExcel12, xlOpenXMLAddIn

$workbook = Get-Item -LiteralPath $Path
$workbookExtension = $workbook.Extension.ToLowerInvariant()
if ($workbook -isnot [IO.FileInfo] -or $workbookExtension -notin '.xlsm', '.xlsb') {
    throw "Expected an .xlsm or .xlsb file: $Path"
}

$sourceDir = Join-Path $workbook.DirectoryName 'vba'
$ribbonPath = Join-Path $workbook.DirectoryName 'ribbon/customUI14.xml'
if (!(Test-Path -LiteralPath $sourceDir -PathType Container)) { throw "Missing VBA directory: $sourceDir" }
$hasRibbon = Test-Path -LiteralPath $ribbonPath -PathType Leaf

if (!$PSBoundParameters.ContainsKey('FileName')) {
    $FileName = if ($workbook.BaseName -eq 'workbook') { $workbook.Directory.Name + $workbookExtension } else { $workbook.Name }
}
$outputExtension = [IO.Path]::GetExtension($FileName).ToLowerInvariant()
if (-not $outputFormats.ContainsKey($outputExtension) -or
    $FileName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) {
    throw "Expected a file name ending in .xlsm, .xlsb or .xlam: $FileName"
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
$updatedWorkbook = Join-Path $workbook.DirectoryName ($publishId + $workbookExtension)
$pendingOutput = [IO.Path]::Combine($outputDirectory, $publishId + $outputExtension)
try {
    # Keep the work book intact until the imports and the release lock succeed.
    [IO.File]::Copy($workbook.FullName, $updatedWorkbook)
    & (Join-Path $PSScriptRoot 'Import-Vba.ps1') -Path $updatedWorkbook -SourceDir $sourceDir | Write-Verbose
    if ($hasRibbon) {
        & (Join-Path $PSScriptRoot 'Add-CustomUI.ps1') -Path $updatedWorkbook -XmlPath $ribbonPath | Write-Verbose
    }

    [void][IO.Directory]::CreateDirectory($outputDirectory)
    Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1')
    $removePersonalInformation = !$KeepPersonalInformation
    if ($removePersonalInformation -or $outputExtension -ne $workbookExtension) {
        # Excel keeps customUI through SaveAs, and the VBA project of the converted file can be locked like that of
        # the work book (docs/notes/publish-addin.md). Lock after the conversion so that the check reads the output.
        # RemovePersonalInformation is stored in the file, so the save of the lock keeps the output clear. It is set
        # only before SaveAs, so the work book is not saved with it (docs/notes/publish-personal-info.md).
        $format = $outputFormats[$outputExtension]
        Use-ExcelWorkbook -Path $updatedWorkbook -ScriptBlock {
            param($xl, $wb)
            if ($removePersonalInformation) { $wb.RemovePersonalInformation = $true }
            $wb.SaveAs($pendingOutput, $format)
        }
    }
    else {
        [IO.File]::Copy($updatedWorkbook, $pendingOutput)
    }
    # The lock reads the locked copy back and names the personal items left in it: its save is the last one and
    # writes the last author unless it is filtered.
    $personalItems = @(& (Join-Path $PSScriptRoot 'Lock-VbaProject.ps1') -Path $pendingOutput)
    if ($personalItems.Count -gt 0) {
        if ($removePersonalInformation) {
            throw ('Personal information remains in the publish copy: ' + ($personalItems -join ', ') +
                '. The publish copy was rejected.')
        }
        Write-Warning ('The output keeps personal information: ' + ($personalItems -join ', ') + '.')
    }
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
