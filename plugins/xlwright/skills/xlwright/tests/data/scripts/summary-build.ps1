param($xl, $wb)
# other.xlsx と data.csv は、試験がブックと同じフォルダに先に作る。
$dir = $wb.Path
$ws = $wb.Worksheets.Item(1)
$ws.Name = 'Data'
$ws.Range('A1').Value2 = 'k'; $ws.Range('B1').Value2 = 'v'
$ws.Range('A2').Value2 = 'x'; $ws.Range('B2').Value2 = [double]1
$ws.Range('A3').Value2 = 'y'; $ws.Range('B3').Value2 = [double]2
$table = $ws.ListObjects.Add(1, $ws.Range('A1:B3'), $null, 1)
$table.Name = 'tblData'
$ws.Range('E10').Formula = "='" + $dir + "\[other.xlsx]Sheet1'!A1"
[void]$wb.Names.Add('Rate', '=Data!$B$2')
[void]$ws.Names.Add('Local', '=Data!$A$2')

$hidden = $wb.Worksheets.Add(); $hidden.Name = 'Hidden'; $hidden.Visible = 0
$veryHidden = $wb.Worksheets.Add(); $veryHidden.Name = 'VeryHidden'; $veryHidden.Visible = 2
$empty = $wb.Worksheets.Add(); $empty.Name = 'Empty'
$chart = $wb.Charts.Add(); $chart.Name = 'Chart1'
# A1 だけに値があるシート。空のシートと UsedRange の番地が同じになる。
$single = $wb.Worksheets.Add(); $single.Name = '単一'
$single.Range('A1').Value2 = '値'

$query = $wb.Worksheets.Add(); $query.Name = 'Query'
$qt = $query.QueryTables.Add('TEXT;' + (Join-Path $dir 'data.csv'), $query.Range('A1'))
$qt.TextFileCommaDelimiter = $true
[void]$qt.Refresh($false)

$module = $wb.VBProject.VBComponents.Add(1)
$module.Name = 'Mod1'
$module.CodeModule.AddFromString("Option Explicit`r`nPublic Sub S()`r`nEnd Sub")
