<#
.SYNOPSIS
Lists the cells of a workbook whose value is an error (#REF!, #NAME?, #DIV/0! and the others).

.DESCRIPTION
Opens the workbook in a hidden Excel with macros and events disabled and without updating links, and closes it
without saving. The workbook is not changed and nothing is recalculated: the values are those Excel holds after
opening, so a workbook saved with manual calculation can show old values (see Calculation).

The used range of every worksheet is read with Range.Value2 in blocks of rows. Range.SpecialCells is not used:
it fails the same way when no cell matches and on a protected sheet, and it took a minute for 140000 scattered
error cells (docs/notes/cell-error.md). Hidden and protected sheets are read too; chart sheets are not.

The result is written as UTF-8 JSON to -OutFile:
- Path.
- Count: the number of error cells found. Truncated: true when Errors lists fewer than Count (-MaxErrors).
- Calculation: "automatic", "manual", "semiautomatic" or "other".
- Sheets, in the order of the workbook: Name, Count, Errors. Each error: Address (e.g. "B3"), Error (e.g.
  "#REF!"), Kind ("formula", "constant", or "spill" for a cell that a formula elsewhere spills into), Formula
  (Range.Formula2 of a formula cell, otherwise null) and SpillParent (the address of the spilling formula, only
  for Kind "spill").
- Error and Warnings.

Exit code is 0 when the workbook was read, also when errors were found, and 1 when it could not be read.

The work is done by Find-ExcelCellError in XlWright.psm1, which a script block can also call to check the cells
in the same Excel instance as other steps.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$OutFile,
    # Only these worksheets, e.g. -Sheet Data,Summary. A name that is not a worksheet of the workbook fails.
    [string[]]$Sheet,
    # Errors lists at most this many cells over all sheets; Count still counts all of them.
    [ValidateRange(0, 1000000)][int]$MaxErrors = 1000,
    # The watchdog stops this Excel instance after this many seconds.
    [ValidateRange(1, 86400)][int]$TimeoutSeconds = 300
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force
# powershell.exe -File passes "a,b" as one string, so split it here.
$Sheet = @($Sheet | ForEach-Object { $_ -split ',' } | Where-Object { $_ })

$result = [ordered]@{
    Path = $null; Count = 0; Truncated = $false; Calculation = $null; Sheets = @(); Error = $null; Warnings = @()
}
$warnings = @()
try {
    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    $result.Path = $full

    # Only plain values leave the block; a COM reference that outlives the workbook keeps Excel from exiting.
    $found = Use-ExcelWorkbook -Path $full -TimeoutSeconds $TimeoutSeconds -WarningVariable +warnings -ScriptBlock {
        param($xl, $wb)
        Find-ExcelCellError -Excel $xl -Workbook $wb -Sheet $Sheet -MaxErrors $MaxErrors
    }
    foreach ($key in @($found.Keys)) { $result[$key] = $found[$key] }
}
catch {
    $result.Error = $_.Exception.Message
}
finally {
    $result.Warnings = @($warnings | ForEach-Object { [string]$_ })
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 6
    [IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}
if ($result.Error) { Write-Error $result.Error; exit 1 }
