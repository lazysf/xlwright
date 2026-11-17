<#
.SYNOPSIS
Reads what a workbook holds (sheets, names, tables, links, connections, VBA components, customUI) in one call.

.DESCRIPTION
Reads the customUI entries from the zip, then opens the workbook in a hidden Excel with macros and events
disabled and without updating links, and closes it without saving. The workbook is not changed.

The result is written as UTF-8 JSON to -OutFile:
- Path.
- Sheets, in the order of the workbook: Name, CodeName, Kind ("worksheet", "chart", "other"), Visible
  ("visible", "hidden", "veryHidden"), UsedRange, Rows, Columns, Protected, Tables (Name, Range).
  UsedRange is null (and Rows and Columns 0) for a sheet without any value or formula and for a sheet
  that is not a worksheet; Excel itself reports A1 for an empty sheet.
- Names: Name, Scope ("workbook" or the name of the sheet), RefersTo, Visible. Names that Excel made,
  e.g. ExternalData_1 of a query, are included.
- ExternalLinks: the workbooks that formulas link to.
- Connections: Name and Type ("oledb", "odbc", "xml", "text", "web", "datafeed", "model", "worksheet",
  "nosource", "other"). The connection string and the command text are left out on purpose: they can
  hold server names and credentials.
- Vba: HasProject, Locked, Components (Name, Kind "standard" / "classes" / "documents" / "forms" / "other",
  Lines), Error. Document modules without code are left out. When the project cannot be read (AccessVBOM
  is not 1, or the project is locked), Components is null and Error says why; the rest is still reported.
- CustomUI: the zip entries under customUI/. Empty when there is none.
- ProtectStructure, Error and Warnings.

Exit code is 0 when the workbook was read, also when only the VBA part could not be.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$OutFile,
    # The watchdog stops this Excel instance after this many seconds.
    [ValidateRange(1, 86400)][int]$TimeoutSeconds = 300
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName Microsoft.VisualBasic

$result = [ordered]@{
    Path = $null; Sheets = @(); Names = @(); ExternalLinks = @(); Connections = @(); Vba = $null; CustomUI = @()
    ProtectStructure = $null; Error = $null; Warnings = @()
}
$warnings = @()
try {
    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    $result.Path = $full

    # Before Excel opens the workbook: while it is open, the zip cannot be read (docs/notes/workbook-summary.md).
    # Only the Open XML formats are zip files.
    if ([IO.Path]::GetExtension($full) -in '.xlsx', '.xlsm', '.xltx', '.xltm', '.xlam') {
        $zip = [IO.Compression.ZipFile]::OpenRead($full)
        try {
            $result.CustomUI = @($zip.Entries | ForEach-Object { $_.FullName } |
                Where-Object { $_ -like 'customUI/*' -and -not $_.EndsWith('/') } | Sort-Object)
        }
        finally { $zip.Dispose() }
    }

    # Only plain values leave the block; a COM reference that outlives the workbook keeps Excel from exiting.
    $found = Use-ExcelWorkbook -Path $full -TimeoutSeconds $TimeoutSeconds -WarningVariable +warnings -ScriptBlock {
        param($xl, $wb)
        $typeName = { param($object) [Microsoft.VisualBasic.Information]::TypeName($object) }

        $visibility = @{ -1 = 'visible'; 0 = 'hidden'; 2 = 'veryHidden' }
        $sheets = @(foreach ($sheet in $wb.Sheets) {
            # Type of a chart sheet is not a sheet type (it was 3, the value of a macro sheet), so the kind
            # comes from the class as well. -4167 is xlWorksheet.
            $class = & $typeName $sheet
            $kind = if ($class -eq 'Worksheet' -and [int]$sheet.Type -eq -4167) { 'worksheet' }
            elseif ($class -eq 'Chart') { 'chart' }
            else { 'other' }
            $row = [ordered]@{
                Name = $sheet.Name; CodeName = $sheet.CodeName; Kind = $kind
                Visible = $visibility[[int]$sheet.Visible]; UsedRange = $null; Rows = 0; Columns = 0
                Protected = [bool]$sheet.ProtectContents; Tables = @()
            }
            if ($kind -eq 'worksheet') {
                # UsedRange of an empty sheet is A1, like that of a sheet with a value only in A1.
                if ($xl.WorksheetFunction.CountA($sheet.Cells) -gt 0) {
                    $used = $sheet.UsedRange
                    $row.UsedRange = $used.Address($false, $false)
                    $row.Rows = $used.Rows.Count
                    $row.Columns = $used.Columns.Count
                }
                $row.Tables = @(foreach ($table in $sheet.ListObjects) {
                    [ordered]@{ Name = $table.Name; Range = $table.Range.Address($false, $false) }
                })
            }
            $row
        })

        $names = @(foreach ($name in $wb.Names) {
            $text = $name.Name
            $scope = 'workbook'
            $parent = $name.Parent
            if ((& $typeName $parent) -ne 'Workbook') {
                # The name of a sheet-level name comes as "<sheet>!<name>".
                $scope = $parent.Name
                $text = $text.Substring($text.LastIndexOf('!') + 1)
            }
            [ordered]@{ Name = $text; Scope = $scope; RefersTo = $name.RefersTo; Visible = [bool]$name.Visible }
        })

        $links = $wb.LinkSources(1)  # xlExcelLinks; null when there is none
        $externalLinks = @(if ($null -ne $links) { foreach ($link in $links) { [string]$link } })

        # XlConnectionType
        $connectionTypes = @{
            1 = 'oledb'; 2 = 'odbc'; 3 = 'xml'; 4 = 'text'; 5 = 'web'; 6 = 'datafeed'; 7 = 'model'; 8 = 'worksheet'
            9 = 'nosource'
        }
        $connections = @(foreach ($connection in $wb.Connections) {
            $type = $connectionTypes[[int]$connection.Type]
            if (-not $type) { $type = 'other' }
            [ordered]@{ Name = $connection.Name; Type = $type }
        })

        $vba = [ordered]@{ HasProject = [bool]$wb.HasVBProject; Locked = $null; Components = $null; Error = $null }
        try {
            $project = $wb.VBProject
            $vba.Locked = ($project.Protection -ne 0)
            if ($vba.Locked) {
                $vba.Error = 'The VBA project is locked; its components cannot be listed.'
            }
            else {
                $componentKinds = @{ 1 = 'standard'; 2 = 'classes'; 3 = 'forms'; 100 = 'documents' }
                $vba.Components = @(foreach ($component in $project.VBComponents) {
                    $lines = $component.CodeModule.CountOfLines
                    if ($component.Type -eq 100 -and $lines -eq 0) { continue }
                    $componentKind = $componentKinds[[int]$component.Type]
                    if (-not $componentKind) { $componentKind = 'other' }
                    [ordered]@{ Name = $component.Name; Kind = $componentKind; Lines = $lines }
                })
            }
        }
        catch {
            # The rest of the summary does not need the VBA project, so this is reported and not thrown.
            $vba.Error = 'The VBA project cannot be read (AccessVBOM must be 1): ' + $_.Exception.Message
        }

        [ordered]@{
            Sheets = $sheets; Names = $names; ExternalLinks = $externalLinks; Connections = $connections; Vba = $vba
            ProtectStructure = [bool]$wb.ProtectStructure
        }
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
