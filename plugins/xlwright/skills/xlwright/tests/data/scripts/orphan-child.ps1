param([string]$Module, [string]$Book, [string]$PidFile)
$ErrorActionPreference = 'Stop'
Import-Module $Module -Force
Use-ExcelWorkbook -Path $Book -TimeoutSeconds 30 -ScriptBlock {
    param($xl, $wb)
    [IO.File]::WriteAllText($PidFile, [string](Get-ExcelProcessId $xl))
    Start-Sleep -Seconds 600
}
