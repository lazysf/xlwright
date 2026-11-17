param($xl, $wb)
$sheet = $wb.Worksheets.Item(1)
Write-ExcelRange -Range $sheet.Range('A1') -Value @(@('日本語ｶﾅ', $null), @($null, $null))
Write-ExcelRange -Range $sheet.Range('B1') -Formula -Value (, @('=1/0'))
@{ Rows = (Read-ExcelRange -Range $sheet.Range('A1:B2')) }
