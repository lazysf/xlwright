param($xl, $wb)
$code = "Public Sub Spin()`r`n    Do`r`n        DoEvents`r`n    Loop`r`nEnd Sub"
$m = $wb.VBProject.VBComponents.Add(1)
$m.CodeModule.AddFromString($code)
[void]$xl.Run('Spin')
