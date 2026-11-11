param($xl, $wb)
$components = $wb.VBProject.VBComponents
$tmp = $components.Add(1)
$tmp.Name = 'TmpSetFont'
$tmp.CodeModule.AddFromString("Public Sub SetFont()`r`n    With ThisWorkbook.VBProject.VBComponents(""frmInfo"").Designer.Controls(""lblInfo"").Font`r`n        .Name = ""Meiryo UI""`r`n        .Size = 11`r`n    End With`r`nEnd Sub")
try { [void]$xl.Run('SetFont') }
finally { $components.Remove($tmp) }
@($components | ForEach-Object { $_.Name })
