# .xlsb をスクリプトに通す

- Date: 2026-11-19
- 出典: 一時フォルダの新しい `.xlsb` に、スキルのスクリプトを順に当てた実測。Excel 16.0.20527.20196（Microsoft 365）、Windows 11、Windows PowerShell 5.1.26100.9818。

`.xlsb` は `.xlsm` と同じ場面（作る、開く・編集・確かめる、配布の作業ブック、配布の出力）で扱う。下の実測で、どのスクリプトも `.xlsm` と同じ手順で通ったため。扱う形式の一覧と理由は、[feature-review.md](./feature-review.md) の「扱うブックの形式」にある。

## 確かめた事実

`New-ExcelWorkbook -Path workbook.xlsb` で作ったブック（`FileFormat` が `50`、`xlExcel12`）に、次を順に当てた。どれも失敗しなかった。

| スクリプト | 結果 |
| --- | --- |
| `Import-Vba.ps1`（`tests/data/vba` の標準・クラス・文書モジュール） | 3 つとも取り込めた |
| `Export-Vba.ps1` | 3 つのファイルと `references.json` を書き出した |
| `Test-VbaCompile.ps1` | `Compiled` が `true` |
| `Invoke-ExcelMacro`（`AddNums 4 5`） | `9` |
| `Write-ExcelRange` と `Read-ExcelRange`（`'001'`、`'=1+1'`） | 文字のまま書けて、読み戻せた |
| `Add-CustomUI.ps1` | `customUI/customUI14.xml` と関連付けを足した |
| `Get-WorkbookSummary.ps1` | `CustomUI` と 3 つの `Vba.Components` を返した |
| `Find-CellError.ps1` | `Count` が `0` |
| `Invoke-RibbonButton.ps1` | 押せて、MsgBox の文言を読めた |
| `Save-WorkbookScreenshot.ps1` | 画像を書いた |
| `Publish-Workbook.ps1`（出力 `.xlsb`、`.xlsm`、`.xlam`） | どれも `VBProject.Protection` が `1`、`FileFormat` が `50`、`52`、`55`、マクロが動き、customUI が残った |
| `Publish-Workbook.ps1`（`-FileName` なし） | `<フォルダ名>.xlsb` を出力した。作業ブックは `FileFormat` が `50`、`Protection` が `0` のまま |
| `Test-RibbonLayout.ps1`（配布した `.xlsb`） | 終了コード 0、食い違い 0 |
| `Publish-Workbook.ps1`（`.xlsm` の作業ブックから `.xlsb` へ） | `FileFormat` が `50`、`Protection` が `1`、マクロが動いた |

UserForm は、試験の `userform` の項目で確かめた。`.xlsm` から書き出したフォーム（`.frm` と `.frx`）を `.xlsb` に取り込むと、デザイン（キャプション、大きさ、コントロール、フレームの中身）は `.xlsm` と同じに読めた。一時の VBA でフォントを変えて保存し、フォームを開くと、初期化の処理とフォントが効いた。`.xlsb` から書き出すと、`.frm` と `.frx` の両方が出た。

`.xlsb` も zip なので、`Add-CustomUI.ps1` と `Get-WorkbookSummary.ps1` は `.xlsm` と同じく zip を直接読み書きする。作業ブックと違う形式で配る時は、`SaveAs` で変換してからロックする（[publish-addin.md](./publish-addin.md)）。

## 未確認

- 大きなブックでの、`.xlsm` との速さやファイルの大きさの差。
