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
Add-Type -AssemblyName Microsoft.VisualBasic
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
        $worksheets = @(foreach ($item in $wb.Sheets) {
            # -4167 is xlWorksheet. The type of a chart sheet is not a sheet type, so the class is checked too.
            if ([Microsoft.VisualBasic.Information]::TypeName($item) -eq 'Worksheet' -and [int]$item.Type -eq -4167) { $item }
        })
        if ($Sheet.Count -gt 0) {
            $known = @($worksheets | ForEach-Object { $_.Name })
            $missing = @($Sheet | Where-Object { $known -notcontains $_ })
            if ($missing.Count -gt 0) {
                throw ('Not a worksheet of the workbook: ' + ($missing -join ', ') + '. Worksheets: ' + ($known -join ', ') + '.')
            }
            $worksheets = @($worksheets | Where-Object { $Sheet -contains $_.Name })
        }

        $calculation = @{ -4105 = 'automatic'; -4135 = 'manual'; 2 = 'semiautomatic' }[[int]$xl.Calculation]
        if (-not $calculation) { $calculation = 'other' }
        $total = 0
        $listed = 0
        $sheets = @(foreach ($worksheet in $worksheets) {
            $used = $worksheet.UsedRange
            $rowCount = $used.Rows.Count
            $columnCount = $used.Columns.Count
            # Bounds the size of one Value2 array also for a wide sheet.
            $blockRows = [Math]::Max(1, [Math]::Min(10000, [Math]::Floor(200000 / $columnCount)))
            $count = 0
            $errors = New-Object System.Collections.Generic.List[object]
            for ($start = 1; $start -le $rowCount; $start += $blockRows) {
                $block = $used.Rows.Item($start).Resize([Math]::Min($blockRows, $rowCount - $start + 1))
                $rows = Read-ExcelRange -Range $block
                for ($r = 0; $r -lt $rows.Count; $r++) {
                    $row = $rows[$r]
                    for ($c = 0; $c -lt $row.Count; $c++) {
                        if ($row[$c] -isnot [Collections.IDictionary]) { continue }
                        $count++
                        if ($listed -ge $MaxErrors) { continue }
                        $listed++
                        $cell = $block.Cells.Item($r + 1, $c + 1)
                        $entry = [ordered]@{
                            Address = $cell.Address($false, $false); Error = $row[$c].Error; Kind = 'constant'
                            Formula = $null; SpillParent = $null
                        }
                        if ($cell.HasFormula) {
                            $entry.Kind = 'formula'
                            $entry.Formula = [string]$cell.Formula2
                        }
                        # SpillParent fails for a cell outside a spill, so HasSpill is checked first.
                        elseif ($cell.HasSpill -eq $true) {
                            $entry.Kind = 'spill'
                            $entry.SpillParent = $cell.SpillParent.Address($false, $false)
                        }
                        $errors.Add($entry)
                    }
                }
            }
            $total += $count
            [ordered]@{ Name = $worksheet.Name; Count = $count; Errors = $errors.ToArray() }
        })
        [ordered]@{ Count = $total; Truncated = ($listed -lt $total); Calculation = $calculation; Sheets = $sheets }
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
