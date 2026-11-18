<#
.SYNOPSIS
Exports workbook VBA modules and UserForms using the repository's layout.

.DESCRIPTION
Standard modules -> <Name>.bas, class modules and document modules (ThisWorkbook, sheets) -> <Name>.cls.
The respective files go into standard/, classes/, and documents/ under OutDir.
UserForms go into forms/ as .frm with the accompanying .frx if exported by VBE.
Document modules without code are skipped.

Text exports are converted from CP932/CRLF to UTF-8 without BOM and LF. Binary .frx files are copied unchanged.

The references of the project (Tools > References) go into references.json in OutDir, in the order of VBProject.References:
{"References": [{Name, Description, Guid, Major, Minor}, ...]}. VBA and Excel (BuiltIn, they cannot be
removed) are left out. A reference to another project (Type 1) is not recorded and its FullPath is listed as a warning.
FullPath is not recorded; it depends on the machine.

Files in OutDir that were not written by this run are listed as a warning (a module removed from the
workbook, or a document module whose code is now empty), because importing them again would bring
the old code back.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$OutDir
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force

$cp932 = [Text.Encoding]::GetEncoding(932, [Text.EncoderFallback]::ExceptionFallback,
    [Text.DecoderFallback]::ExceptionFallback)
$utf8 = New-Object System.Text.UTF8Encoding($false)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path -LiteralPath $OutDir).ProviderPath
$folders = @{ 1 = 'standard'; 2 = 'classes'; 3 = 'forms'; 100 = 'documents' }
if (Get-ChildItem -LiteralPath $OutDir -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm', '.frx' }) {
    throw "Place VBA files in standard/, classes/, documents/, or forms/, not directly in $OutDir"
}
$outputDirectories = @(foreach ($folder in $folders.Values) {
    $directory = Join-Path $OutDir $folder
    New-Item -ItemType Directory -Force -Path $directory | Out-Null
    $directory
})
$tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('vba-export-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $tmpDir | Out-Null

try {
    $written = Use-ExcelWorkbook -Path $Path -ScriptBlock {
        param($xl, $wb)
        foreach ($c in $wb.VBProject.VBComponents) {
            if ($c.Type -eq 1) { $ext = 'bas' }
            elseif ($c.Type -eq 2) { $ext = 'cls' }
            elseif ($c.Type -eq 3) { $ext = 'frm' }
            elseif ($c.Type -eq 100) {
                if ($c.CodeModule.CountOfLines -eq 0) { continue }
                $ext = 'cls'
            }
            else {
                Write-Warning ("Skipped " + $c.Name + " (component type " + $c.Type + " is not supported)")
                continue
            }
            $file = $c.Name + '.' + $ext
            $tmp = Join-Path $tmpDir $file
            $c.Export($tmp)
            try { $text = $cp932.GetString([IO.File]::ReadAllBytes($tmp)) }
            catch { throw ("Cannot decode Excel export as CP932: " + $file) }
            $text = $text -replace "`r`n?", "`n"
            $directory = Join-Path $OutDir $folders[[int]$c.Type]
            $outputFile = Join-Path $directory $file
            if ($ext -eq 'frm') {
                $sidecarName = $c.Name + '.frx'
                $sidecar = Join-Path $tmpDir $sidecarName
                if (Test-Path -LiteralPath $sidecar -PathType Leaf) {
                    $outputSidecar = Join-Path $directory $sidecarName
                    [IO.File]::Copy($sidecar, $outputSidecar, $true)
                    $outputSidecar
                }
            }
            [IO.File]::WriteAllText($outputFile, $text, $utf8)
            $outputFile
        }
        $references = @(foreach ($r in $wb.VBProject.References) {
            if ($r.BuiltIn) { continue }
            if ($r.Type -ne 0) {
                # The workbook is open with macros disabled, so the referenced project is not loaded and only FullPath can be read.
                Write-Warning ("Not recorded in references.json: " + $r.FullPath + " (a reference to another project)")
                continue
            }
            [ordered]@{
                Name = [string]$r.Name; Description = [string]$r.Description; Guid = [string]$r.Guid
                Major = [int]$r.Major; Minor = [int]$r.Minor
            }
        })
        $json = (ConvertTo-Json -InputObject ([ordered]@{ References = $references }) -Depth 3) -replace "`r`n", "`n"
        $referencesFile = Join-Path $OutDir 'references.json'
        [IO.File]::WriteAllText($referencesFile, $json + "`n", $utf8)
        $referencesFile
    }
}
finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

$written | ForEach-Object { "exported: " + $_.Substring($OutDir.Length).TrimStart('\', '/') }
$stale = @(Get-ChildItem -LiteralPath $outputDirectories -File |
    Where-Object { $_.Extension -in '.bas', '.cls', '.frm', '.frx' -and $written -notcontains $_.FullName })
if ($stale.Count -gt 0) {
    Write-Warning ("Present in OutDir but not written by this export (module removed, or document module now empty): " +
        (($stale | ForEach-Object { $_.FullName.Substring($OutDir.Length).TrimStart('\', '/') }) -join ', '))
}
