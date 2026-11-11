param($xl, $wb)
$wb.Worksheets.Item(1).Range('A1').Value2 = '別ブックの値'
$xl.Run('Probe')
