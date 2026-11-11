param($xl, $wb)
$ws = $wb.Worksheets.Item('集計表')
@{ A1 = $ws.Range('A1').Value2; B1 = $ws.Range('B1').Formula; B1Value = $ws.Range('B1').Value2; Rows = $ws.Range('A1:B2').Value2 }
$ws
@($wb.Worksheets | ForEach-Object { $_.Name })
