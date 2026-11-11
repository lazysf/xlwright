<#
.SYNOPSIS
Runs a script file against a workbook through Use-ExcelWorkbook and writes the result as UTF-8 JSON.

.DESCRIPTION
The script file holds the body of a script block that starts with param($xl, $wb). It is read as UTF-8
(with or without BOM), so it may contain Japanese text. Whatever the block outputs becomes Result in the
JSON file. Return plain values (strings, numbers, booleans, hashtables, arrays); COM objects are replaced
with the text "[COM object]". A two-dimensional array (Range.Value2 of several cells) becomes an array of rows.

The JSON has Ok, Result, Error and Warnings. The exit code is 1 when the block or Excel failed.
With -Create the workbook is first created with New-ExcelWorkbook (it must not exist yet).
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$Script,
    [Parameter(Mandatory)][string]$OutFile,
    [switch]$Create,
    [switch]$Save,
    [switch]$Visible,
    [switch]$EnableMacros,
    [switch]$EnableEvents,
    [ValidateRange(1, 86400)][int]$TimeoutSeconds = 300
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force

$utf8 = New-Object System.Text.UTF8Encoding($false)
$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)

function ConvertTo-PlainValue($Value, [int]$Depth) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.__ComObject]) { return '[COM object]' }
    if ($Depth -ge 6) { return [string]$Value }
    if ($Value -is [string] -or $Value -is [ValueType]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $map = [ordered]@{}
        foreach ($key in $Value.Keys) { $map[[string]$key] = ConvertTo-PlainValue $Value[$key] ($Depth + 1) }
        return $map
    }
    if ($Value -is [Array] -and $Value.Rank -eq 2) {
        # Range.Value2 of several cells: a 1-based [row, column] array. Keep it as rows.
        return , @(for ($r = $Value.GetLowerBound(0); $r -le $Value.GetUpperBound(0); $r++) {
            , @(for ($c = $Value.GetLowerBound(1); $c -le $Value.GetUpperBound(1); $c++) {
                ConvertTo-PlainValue $Value[$r, $c] ($Depth + 1)
            })
        })
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        return , @(foreach ($item in $Value) { ConvertTo-PlainValue $item ($Depth + 1) })
    }
    [string]$Value
}

$result = [ordered]@{ Ok = $false; Result = $null; Error = $null; Warnings = @() }
$warnings = @()
try {
    $text = $strictUtf8.GetString([IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Script).ProviderPath))
    $block = [scriptblock]::Create($text.TrimStart([char]0xFEFF))
    if ($Create) {
        [void](New-ExcelWorkbook -Path $Path -TimeoutSeconds ([Math]::Min($TimeoutSeconds, 120)) -WarningVariable +warnings)
    }
    # Convert inside the block so that no COM reference outlives the workbook; a held reference keeps
    # Excel from exiting after Quit.
    $wrapper = {
        param($xl, $wb)
        $output = & $block $xl $wb
        , @(foreach ($item in @($output)) { ConvertTo-PlainValue $item 0 })
    }
    $converted = Use-ExcelWorkbook -Path $Path -ScriptBlock $wrapper -Save:$Save -Visible:$Visible `
        -EnableMacros:$EnableMacros -EnableEvents:$EnableEvents -TimeoutSeconds $TimeoutSeconds `
        -WarningVariable +warnings
    $result.Result = @($converted)
    $result.Ok = $true
}
catch {
    $result.Error = $_.Exception.Message
}
finally {
    $result.Warnings = @($warnings | ForEach-Object { [string]$_ })
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 10
    [IO.File]::WriteAllText($OutFile, $json, $utf8)
}
if (-not $result.Ok) { exit 1 }
