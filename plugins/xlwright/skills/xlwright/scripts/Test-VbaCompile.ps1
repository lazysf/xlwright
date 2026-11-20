<#
.SYNOPSIS
Compiles the VBA project of a workbook and reports the compile error, if any.

.DESCRIPTION
Opens the workbook in a hidden Excel with macros disabled, makes its project the active one in the VBE and
executes the VBE command "Compile VBAProject" (command bar control 578). There is no COM method for
compiling. The workbook is closed without saving.

The outcome is read from what the VBE does next:
- The control becomes disabled: the project compiled.
- A dialog (#32770) of this Excel instance appears: a compile error. Its text and the selection of the
  active code pane (the place of the error) are recorded, the dialog is closed with WM_CLOSE and the VBE
  window, which the error makes visible on the screen, is hidden again.
- Neither within -DialogTimeoutSeconds: reported as an error, not as a pass.

Only one compile error is reported: the one the VBE shows. The place is left null when the VBE has no
active code pane. A locked project is not compiled and is reported as an error.

If the dialog cannot be closed, the Excel instance started by this script is stopped by its PID and start
time. Other Excel processes are never touched.

The work is done by Test-ExcelVbaCompile in XlWright.psm1, which a script block can also call to import, compile
and run macros in one Excel instance.

The result is written as UTF-8 JSON to -OutFile (stdout of powershell.exe garbles Japanese in Git Bash):
Compiled (true / false, null when the check itself failed), CompileError (Message, Module, Line, Column)
and Error. Exit code is 0 only when the project compiled.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$OutFile,
    # The watchdog stops this Excel instance after this many seconds.
    [ValidateRange(1, 86400)][int]$TimeoutSeconds = 300,
    # How long to wait for the compile to finish or the error dialog to appear.
    [ValidateRange(1, 3600)][int]$DialogTimeoutSeconds = 5
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'XlWright.psm1') -Force

$result = [ordered]@{ Compiled = $null; CompileError = $null; Error = $null }
try {
    $checked = Use-ExcelWorkbook -Path $Path -TimeoutSeconds $TimeoutSeconds -ScriptBlock {
        param($xl, $wb)
        Test-ExcelVbaCompile -Excel $xl -Workbook $wb -DialogTimeoutSeconds $DialogTimeoutSeconds
    }
    $result.Compiled = $checked.Compiled
    $result.CompileError = $checked.CompileError
}
catch {
    # A compile error whose dialog could not be closed: what was read is kept next to the error.
    $partial = $_.Exception.Data['CompileResult']
    if ($null -ne $partial) {
        $result.Compiled = $partial.Compiled
        $result.CompileError = $partial.CompileError
    }
    $result.Error = $_.Exception.Message
}
finally {
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 5
    [IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}
if ($result.Error) { Write-Error $result.Error; exit 1 }
if (-not $result.Compiled) { exit 1 }
