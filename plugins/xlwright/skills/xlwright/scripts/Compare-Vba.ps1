<#
.SYNOPSIS
Compares the VBA of a workbook with the source files and reports the components that differ.

.DESCRIPTION
Exports the workbook with Export-Vba.ps1 into a temporary folder and compares the files with those in
standard/, classes/, documents/ and forms/ under -SourceDir. Neither the workbook nor the sources are changed.

Before comparing, only this is levelled out: line endings, blank lines at the end of a file, and blank lines
right after the Attribute lines (the VBE adds one to a .frm on every import). Everything else counts,
e.g. the spelling of keywords and spaces the VBE rewrites on import: export once to bring the sources into
the form of the VBE.

Status of a component:
- same: no difference.
- case_only: the texts differ only in upper and lower case. The VBE unifies the case of an identifier in
  the whole project, so this is often not an edit; it cannot be told from a changed string literal.
- different: any other difference.
- source_only / workbook_only: present on one side only. A document module without code is not exported,
  so it is never workbook_only.

FirstDifference holds the first differing line: Line (the line number in the source file), Source and
Workbook (null where that side has no such line).

The .frx of a UserForm is not compared; its bytes differ between two exports of the same workbook. Only
the .frm text is.

The result is written as UTF-8 JSON to -OutFile: InSync, Components (Name, Kind, Status, FirstDifference),
Error and Warnings. Exit code is 0 only when every component is the same.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$SourceDir,
    [Parameter(Mandatory)][string]$OutFile
)
$ErrorActionPreference = 'Stop'

$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
$extensions = [ordered]@{ standard = '.bas'; classes = '.cls'; documents = '.cls'; forms = '.frm' }

# The lines that are compared, each with its number in the file.
function Get-ComparedLine([string]$File) {
    try { $text = $strictUtf8.GetString([IO.File]::ReadAllBytes($File)) }
    catch { throw ("Invalid UTF-8 in " + $File) }
    $lines = $text.TrimStart([char]0xFEFF) -split "\r\n|\r|\n"
    $last = $lines.Count - 1
    while ($last -ge 0 -and $lines[$last] -match '^\s*$') { $last-- }
    # The header ends with the Attribute lines that start at VB_Name; the code follows.
    $code = 0
    for ($i = 0; $i -le $last; $i++) {
        if ($lines[$i] -like 'Attribute VB_Name = *') {
            $code = $i + 1
            while ($code -le $last -and $lines[$code] -like 'Attribute *') { $code++ }
            break
        }
    }
    $skipTo = $code
    if ($code -gt 0) { while ($skipTo -le $last -and $lines[$skipTo] -match '^\s*$') { $skipTo++ } }
    for ($i = 0; $i -le $last; $i++) {
        if ($i -ge $code -and $i -lt $skipTo) { continue }
        [pscustomobject]@{ Number = $i + 1; Text = $lines[$i] }
    }
}

function Compare-Component([string]$SourceFile, [string]$WorkbookFile) {
    $source = @(Get-ComparedLine $SourceFile)
    $workbook = @(Get-ComparedLine $WorkbookFile)
    $first = $null
    $caseOnly = $source.Count -eq $workbook.Count
    for ($i = 0; $i -lt [Math]::Max($source.Count, $workbook.Count); $i++) {
        $s = if ($i -lt $source.Count) { $source[$i].Text } else { $null }
        $w = if ($i -lt $workbook.Count) { $workbook[$i].Text } else { $null }
        # A missing line is not an empty one; PowerShell would pass $null to Equals as "".
        if ($null -eq $s -or $null -eq $w) {
            $caseOnly = $false
            if ($null -eq $first) {
                $line = if ($null -ne $s) { $source[$i].Number } elseif ($source.Count -gt 0) { $source[-1].Number + 1 } else { 1 }
                $first = [ordered]@{ Line = $line; Source = $s; Workbook = $w }
            }
            break
        }
        if ([string]::Equals($s, $w, [StringComparison]::Ordinal)) { continue }
        if ($null -eq $first) { $first = [ordered]@{ Line = $source[$i].Number; Source = $s; Workbook = $w } }
        if (-not [string]::Equals($s, $w, [StringComparison]::OrdinalIgnoreCase)) { $caseOnly = $false; break }
    }
    $status = if ($null -eq $first) { 'same' } elseif ($caseOnly) { 'case_only' } else { 'different' }
    [ordered]@{ Status = $status; FirstDifference = $first }
}

function Get-ComponentFile([string]$Root) {
    $found = @{}
    foreach ($kind in $extensions.Keys) {
        $directory = Join-Path $Root $kind
        if (-not (Test-Path -LiteralPath $directory -PathType Container)) { continue }
        foreach ($file in Get-ChildItem -LiteralPath $directory -File) {
            if ($file.Extension -ne $extensions[$kind]) { continue }
            # VBA names are not case-sensitive.
            $found[($kind + '/' + $file.BaseName).ToLowerInvariant()] = [pscustomobject]@{ Kind = $kind; Name = $file.BaseName; File = $file.FullName }
        }
    }
    $found
}

$components = New-Object System.Collections.Generic.List[object]
$warnings = @()
$errorText = $null
$tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('vba-compare-' + [guid]::NewGuid())
try {
    if (-not (Test-Path -LiteralPath $SourceDir -PathType Container)) { throw "Source directory not found: $SourceDir" }
    $SourceDir = (Resolve-Path -LiteralPath $SourceDir).ProviderPath
    New-Item -ItemType Directory -Path $tmpDir | Out-Null
    # The export is the one way this skill reads VBA out of a workbook; it opens the workbook without saving.
    $exportOutput = & (Join-Path $PSScriptRoot 'Export-Vba.ps1') -Path $Path -OutDir $tmpDir 3>&1
    $warnings = @($exportOutput | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } | ForEach-Object { $_.Message })

    $sourceFiles = Get-ComponentFile $SourceDir
    $workbookFiles = Get-ComponentFile $tmpDir
    $keys = @($sourceFiles.Keys) + @($workbookFiles.Keys) | Sort-Object -Unique
    $kindOrder = @($extensions.Keys)
    $rows = foreach ($key in $keys) {
        $s = $sourceFiles[$key]
        $w = $workbookFiles[$key]
        $compared = if ($null -eq $w) { [ordered]@{ Status = 'source_only'; FirstDifference = $null } }
        elseif ($null -eq $s) { [ordered]@{ Status = 'workbook_only'; FirstDifference = $null } }
        else { Compare-Component $s.File $w.File }
        $known = if ($null -ne $s) { $s } else { $w }
        [pscustomobject]@{
            Name = $known.Name; Kind = $known.Kind; Status = $compared.Status; FirstDifference = $compared.FirstDifference
            Order = $kindOrder.IndexOf($known.Kind)
        }
    }
    foreach ($row in @($rows | Sort-Object Order, Name)) {
        $components.Add([ordered]@{ Name = $row.Name; Kind = $row.Kind; Status = $row.Status; FirstDifference = $row.FirstDifference })
    }
}
catch {
    $errorText = $_.Exception.Message
}
finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
    $inSync = (-not $errorText) -and @($components | Where-Object { $_.Status -ne 'same' }).Count -eq 0
    $result = [ordered]@{
        InSync = $inSync; Components = $components.ToArray(); Error = $errorText
        Warnings = @($warnings | ForEach-Object { [string]$_ })
    }
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 5
    [IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}
if ($errorText) { Write-Error $errorText; exit 1 }
if (-not $inSync) { exit 1 }
