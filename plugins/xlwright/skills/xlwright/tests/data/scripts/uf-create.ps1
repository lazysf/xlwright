param($xl, $wb)
$c = $wb.VBProject.VBComponents.Add(3)
$c.Name = 'frmInfo'
$c.Properties.Item('Caption').Value = '情報フォーム'
$c.Properties.Item('Width').Value = 240
$c.Properties.Item('Height').Value = 160
$d = $c.Designer
$lbl = $d.Controls.Add('Forms.Label.1', 'lblInfo')
$lbl.Caption = '初期値'
$lbl.Left = 12; $lbl.Top = 12; $lbl.Width = 200; $lbl.Height = 18
$fra = $d.Controls.Add('Forms.Frame.1', 'fraOpt')
$fra.Caption = 'オプション'
$fra.Left = 12; $fra.Top = 36; $fra.Width = 200; $fra.Height = 50
$chk = $fra.Controls.Add('Forms.CheckBox.1', 'chkShow')
$chk.Caption = '表示する'
$chk.Left = 6; $chk.Top = 6; $chk.Width = 120; $chk.Height = 18
$btn = $d.Controls.Add('Forms.CommandButton.1', 'btnOk')
$btn.Caption = 'OK'
$btn.Left = 150; $btn.Top = 96; $btn.Width = 60; $btn.Height = 24

$code = "Option Explicit`r`n`r`n" +
    "Private Sub UserForm_Initialize()`r`n" +
    "    lblInfo.Caption = ThisWorkbook.Worksheets(1).Range(""A1"").Text`r`n" +
    "    chkShow.Value = True`r`n" +
    "End Sub`r`n`r`n" +
    "Private Sub btnOk_Click()`r`n" +
    "    Unload Me`r`n" +
    "End Sub"
$c.CodeModule.AddFromString($code)

$p = $wb.VBProject.VBComponents.Add(1)
$p.Name = 'FormProbe'
$p.CodeModule.AddFromString("Option Explicit`r`n`r`nPublic Function Probe() As String`r`n    Load frmInfo`r`n    Probe = frmInfo.Caption & ""|"" & frmInfo.lblInfo.Caption & ""|"" & frmInfo.chkShow.Value & ""|"" & frmInfo.lblInfo.Font.Name`r`n    Unload frmInfo`r`nEnd Function")

$font = 'not raised'
try { $lbl.Font.Name = 'Meiryo UI' } catch { $font = $_.Exception.GetType().Name }
@{ FontAssignment = $font }
