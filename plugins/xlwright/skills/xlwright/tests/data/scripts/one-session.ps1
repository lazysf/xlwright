param($xl, $wb)
# ブックの隣の vba/ を取り込み、コンパイルし、マクロを動かし、エラー値のセルを探す。すべて 1 つの Excel で行う
$imported = @(Import-ExcelVba -Workbook $wb -SourceDir (Join-Path $wb.Path 'vba'))
$compile = Test-ExcelVbaCompile -Excel $xl -Workbook $wb
$ws = $wb.Worksheets.Item(1)
$ws.Range('A1').Formula = '=1/0'
@{
    Imported = $imported
    Compiled = $compile.Compiled
    CompileError = $compile.CompileError
    Sum = Invoke-ExcelMacro -Excel $xl -Name 'Util.AddNums' -ArgumentList 2, 3
    Errors = Find-ExcelCellError -Excel $xl -Workbook $wb
}
