<#
.SYNOPSIS
Checks what this skill needs before a workbook is opened and reports what stands in the way.

.DESCRIPTION
Excel is not started, and no setting, workbook or Git state is changed. The one thing that is done is to start
a hidden "cmd.exe /c exit" through WMI, because only that shows whether the watchdog can be started.

Checked: a running Excel, AccessVBOM, WMI, and with -Path the workbook (exists, open somewhere, Git state).
A problem keeps Ready false; see Get-ExcelPrerequisite in XlWright.psm1 for what counts as one. AccessVBOM is
a problem only with -Vba. The Git state never is.

The result is written as UTF-8 JSON to -OutFile: Ready, Problems (one sentence each, with what to do), Excel
(Running, Processes: Id, StartTime, WindowTitle), AccessVBOM (Value, Enabled), Wmi (Ok, Error), Workbook (null
without -Path; Path, Exists, OpenElsewhere, Git: Available, InRepository, Tracked, Uncommitted) and Error.
Exit code is 0 only when Ready is true.
#>
param(
    [Parameter(Mandatory)][string]$OutFile,
    # The workbook that is going to be opened.
    [string]$Path,
    # The work reads or changes VBA, which needs AccessVBOM.
    [switch]$Vba
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force

$result = [ordered]@{
    Ready = $false; Problems = @(); Excel = $null; AccessVBOM = $null; Wmi = $null; Workbook = $null; Error = $null
}
try {
    $found = Get-ExcelPrerequisite -Path $Path -Vba:$Vba
    foreach ($key in @($found.Keys)) { $result[$key] = $found[$key] }
}
catch {
    $result.Error = $_.Exception.Message
}
finally {
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 6
    [IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}
if ($result.Error) { Write-Error $result.Error; exit 1 }
if (-not $result.Ready) { exit 1 }
