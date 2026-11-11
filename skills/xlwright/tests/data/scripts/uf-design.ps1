param($xl, $wb)
$c = $wb.VBProject.VBComponents.Item('frmInfo')
$d = $c.Designer
$controls = foreach ($x in $d.Controls) {
    [ordered]@{ Name = $x.Name; Caption = $x.Caption; Left = $x.Left; Top = $x.Top; Width = $x.Width;
        Height = $x.Height; TabIndex = $x.TabIndex; Parent = $x.Parent.Name }
}
[ordered]@{
    Type = $c.Type
    Caption = $c.Properties.Item('Caption').Value
    Width = $c.Properties.Item('Width').Value
    FrameChildren = @($d.Controls.Item('fraOpt').Controls | ForEach-Object { $_.Name })
    Controls = @($controls)
}
