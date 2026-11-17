param([string]$Module, [string]$Book, [string]$PidFile)
$ErrorActionPreference = 'Stop'
Import-Module $Module -Force
# The limit has to outlast the start of Excel, the PID file, the kill of this process and the 2 s the test then waits.
Use-ExcelWorkbook -Path $Book -TimeoutSeconds 12 -ScriptBlock {
    param($xl, $wb)
    [IO.File]::WriteAllText($PidFile, [string](Get-ExcelProcessId $xl))
    Start-Sleep -Seconds 600
}
