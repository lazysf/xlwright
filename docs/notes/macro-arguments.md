# Invoke-ExcelMacro に渡す引数

- Date: 2026-11-20
- 出典: 一時フォルダの新しい `.xlsm` で、実験用スクリプトを非表示の Excel で流した実測。Microsoft 365 Excel 64-bit 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818。1 項目ごとに別の `powershell.exe` で流した。

`Invoke-ExcelMacro` は、`-ArgumentList` の要素が PSObject で包まれていれば、中の値を `Application.Run` に渡す。包んだまま渡すと、`String` の引数で `DISP_E_TYPEMISMATCH`（`0x80020005`）になるため。実装は `plugins/xlwright/skills/xlwright/scripts/XlWright.psm1` の `Invoke-ExcelMacro` を正本とする。

## 確かめた事実

標準モジュールに `Function Probe(wb As Workbook, a As String, b As String) As Long` を、Private と Public で置いた。呼び方 2 つ、Private と Public、文字列の渡し方 4 つ、Workbook の渡し方 2 つで、32 通りを 1 回ずつ流した。

| 文字列の渡し方 | `-is [psobject]` | 包みを外す前の `Invoke-ExcelMacro` | `$xl.Run(...)` |
| --- | --- | --- | --- |
| literal（`'abc'`） | `False` | 通る | 通る |
| `[psobject]'abc'` | `True` | `0x80020005` | 通る |
| `Join-Path 'a' 'b'` の出力 | `True` | `0x80020005` | 通る |
| `Get-Content` の 1 行 | `True` | `0x80020005` | 通る |

- Private か Public かでは、結果は変わらなかった。
- 包みを外す前の `Invoke-ExcelMacro` は、引数を `[System.__ComObject].InvokeMember('Run', ...)` にそのまま渡していた。`$xl.Run(...)` は PowerShell の COM の束縛を通り、包んだ値も通った。
- `[psobject]$wb` と書いても、Workbook は `-is [psobject]` が `False` のままだった。PSObject で包んだ Workbook を渡す場合は、確かめていない。
- 包みを外した後は、32 通りがすべて通った。`Invoke-ExcelMacro` の値は、`$xl.Run` と同じだった。
- セルへの代入でも、PSObject で包んだ値は `InvokeMember` で落ちた（[range-value.md](./range-value.md)）。`Write-ExcelRange` も包みを外してから書く。

外部の実機調査で、アドインの Private 関数を `Workbook, String, String` の引数で呼ぶと、`Invoke-ExcelMacro` だけが `0x80020005` で落ちたと報告があった。上の実測でも、同じ HRESULT が出て、`$xl.Run` では通るという同じ違いが出た。報告の事象の原因がこれかは、報告のあったアドインでは確かめていない。

`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` の `macro` の項目が、包んだ文字列を Private 関数の `String` の引数に渡す呼び出しを確かめる（2026-11-20）。包みを外さない写しでは、この項目が `0x80020005` で落ちた。

## 未確認

- PSObject で包んだ Workbook などの COM のオブジェクトを渡した時の振る舞い。
- 配列の引数で、要素が PSObject で包まれている場合。`Invoke-ExcelMacro` は配列の中までは外さない。
- PowerShell 7 での振る舞い。
