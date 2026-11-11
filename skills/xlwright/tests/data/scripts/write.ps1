param($xl, $wb)
$ws = $wb.Worksheets.Item(1)
$ws.Name = '集計表'
$ws.Range('A1').Value2 = [double]2
$ws.Range('B1').Formula = '=A1*3'
$ws.Range('A2').Value2 = '日本語ｶﾅ'
'書き込み完了'
