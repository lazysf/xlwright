<#
.SYNOPSIS
Imports UTF-8 (no BOM), LF .bas / .cls / .frm files and saves the workbook.

.DESCRIPTION
- All sources are checked for lossless CP932 conversion before opening the workbook.
- Temporary files use CP932 and CRLF. VBComponents.Import misreads an LF-only .cls as a standard module.
- An existing standard module, class module, or UserForm with the same name and type is replaced.
- A UserForm's referenced .frx is copied unchanged next to the temporary .frm before import.
- A document module (ThisWorkbook, sheets) cannot be imported. Its code is replaced instead:
  the class header and the leading Attribute lines are dropped and the rest is written to the
  existing module. Attribute lines inside procedures are not handled.
- The standard/, classes/, documents/, and forms/ directories determine the component type.
- Every existing component must match the source type before any code is changed.
- After import, each module's name and type are checked against the file.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$SourceDir
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force

$cp932 = [Text.Encoding]::GetEncoding(932, [Text.EncoderFallback]::ExceptionFallback,
    [Text.DecoderFallback]::ExceptionFallback)
if (Get-ChildItem -LiteralPath $SourceDir -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm', '.frx' }) {
    throw "Place VBA files in standard/, classes/, documents/, or forms/, not directly in $SourceDir"
}
$folderTypes = @{ standard = 1; classes = 2; forms = 3; documents = 100 }
$extensions = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm'; 100 = '.cls' }
$sourceFolders = @(Get-ChildItem -LiteralPath $SourceDir -Directory |
    Where-Object { $folderTypes.ContainsKey($_.Name) })
$files = @(foreach ($folder in $sourceFolders) {
    Get-ChildItem -LiteralPath $folder.FullName -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm' }
})
if ($files.Count -eq 0) { throw "No .bas / .cls / .frm files in VBA component directories under $SourceDir" }

$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
$sources = foreach ($f in $files) {
    $bytes = [IO.File]::ReadAllBytes($f.FullName)
    try { $text = $strictUtf8.GetString($bytes) }
    catch { throw ("Invalid UTF-8 in " + $f.FullName) }
    if ($text.StartsWith([string][char]0xFEFF, [StringComparison]::Ordinal) -or $text.Contains("`r")) {
        throw ("Expected UTF-8 without BOM and LF line endings in " + $f.FullName)
    }
    $text = $text.Replace("`n", "`r`n")
    try { $importBytes = $cp932.GetBytes($text) }
    catch { throw ("Cannot encode as CP932: " + $f.FullName) }
    if (![string]::Equals($cp932.GetString($importBytes), $text, [StringComparison]::Ordinal)) {
        throw ("Cannot round-trip through CP932: " + $f.FullName)
    }
    $m = [regex]::Match($text, '(?m)^Attribute VB_Name = "([^"]+)"\r$')
    if (-not $m.Success) { throw ("No Attribute VB_Name in " + $f.Name) }
    $name = $m.Groups[1].Value
    $type = $folderTypes[$f.Directory.Name]
    if ($f.Extension -ne $extensions[$type] -or $f.BaseName -ne $name) {
        throw ("Unexpected extension or VB_Name in " + $f.FullName)
    }
    $sidecar = $null
    if ($type -eq 3) {
        $designer = $text.Substring(0, $m.Index)
        if ($designer -match '(?m)^[ \t]*OleObjectBlob[ \t]*=') {
            $blob = [regex]::Match($designer, '(?m)^[ \t]*OleObjectBlob[ \t]*=[ \t]*"([^"]+)":[0-9A-Fa-f]+[ \t]*\r$')
            $sidecarName = $f.BaseName + '.frx'
            if (!$blob.Success -or $blob.Groups[1].Value -ne $sidecarName) {
                throw ("Expected a reference to " + $sidecarName + " in " + $f.FullName)
            }
            $sidecar = Join-Path $f.DirectoryName $sidecarName
            if (!(Test-Path -LiteralPath $sidecar -PathType Leaf)) {
                throw ("Missing UserForm resource: " + $sidecar)
            }
        }
    }
    [pscustomobject]@{
        File = $f.Name
        Name = $name
        Text = $text
        ImportBytes = $importBytes
        Sidecar = $sidecar
        Type = $type
    }
}

$duplicate = $sources | Group-Object Name | Where-Object Count -gt 1
if ($duplicate) { throw ("Duplicate VB_Name: " + ($duplicate.Name -join ', ')) }

$tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('vba-import-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $tmpDir | Out-Null
try {
    Use-ExcelWorkbook -Path $Path -Save -ScriptBlock {
        param($xl, $wb)
        $components = $wb.VBProject.VBComponents
        $existingByName = @{}
        foreach ($component in $components) { $existingByName[$component.Name] = $component }
        foreach ($s in $sources) {
            $existing = $existingByName[$s.Name]
            if ($null -ne $existing -and $existing.Type -ne $s.Type) {
                throw ("Component type mismatch for " + $s.Name + ": workbook " + $existing.Type + ", source " + $s.Type)
            }
            if ($s.Type -eq 100 -and $null -eq $existing) {
                throw ("No document module named " + $s.Name + " in the workbook")
            }
        }
        foreach ($s in $sources) {
            $existing = $existingByName[$s.Name]
            if ($s.Type -eq 100) {
                $lines = $s.Text -split "`r`n"
                $i = 0
                if ($lines[0] -eq 'VERSION 1.0 CLASS') {
                    while ($i -lt $lines.Count -and $lines[$i] -ne 'END') { $i++ }
                    $i++
                }
                while ($i -lt $lines.Count -and $lines[$i] -like 'Attribute *') { $i++ }
                $code = ''
                if ($i -lt $lines.Count) { $code = ($lines[$i..($lines.Count - 1)] -join "`r`n").TrimEnd("`r", "`n") }
                $cm = $existing.CodeModule
                if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
                if ($code.Length -gt 0) { $cm.AddFromString($code) }
                "replaced document module: " + $s.Name
                continue
            }
            $tmp = Join-Path $tmpDir $s.File
            [IO.File]::WriteAllBytes($tmp, $s.ImportBytes)
            if ($s.Sidecar) {
                [IO.File]::Copy($s.Sidecar, [IO.Path]::ChangeExtension($tmp, '.frx'))
            }
            if ($null -ne $existing) { $components.Remove($existing) }
            $c = $components.Import($tmp)
            if ($c.Name -ne $s.Name -or $c.Type -ne $s.Type) {
                throw ("Imported " + $s.File + " as " + $c.Name + " (type " + $c.Type + "), expected " +
                    $s.Name + " (type " + $s.Type + ")")
            }
            "imported: " + $s.Name
        }
    }
}
finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}
