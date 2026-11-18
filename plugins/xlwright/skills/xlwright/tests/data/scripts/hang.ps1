param($xl, $wb)
$code = "Public Sub Spin()`r`n    Do`r`n        DoEvents`r`n    Loop`r`nEnd Sub"
$m = $wb.VBProject.VBComponents.Add(1)
$m.CodeModule.AddFromString($code)
# Through Invoke-ExcelMacro, so that its failure path is covered when the watchdog stops Excel.
Invoke-ExcelMacro -Excel $xl -Name ($m.Name + '.Spin')
