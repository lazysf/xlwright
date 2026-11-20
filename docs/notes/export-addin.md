# 配布の出力をアドイン（.xlam）にする

- Date: 2026-11-19
- 出典: 一時フォルダの新しいブックで `Import-Vba.ps1`、`Add-CustomUI.ps1`、`Lock-VbaProject.ps1`、`Test-RibbonLayout.ps1` を流した実測と、`Test-XlWright.ps1 -Full -Case publish`（今の `export`）。Excel 16.0.20527.20196（Microsoft 365）、Windows 11、Windows PowerShell 5.1.26100.9818。

`Export-Workbook.ps1` は、作業ブックを `.xlsm`（か `.xlsb`）のまま扱い、出力の時だけ `.xlam` に変える。アドインは作業中にシートを見せないので、開発と確かめる作業はブックの方が扱いやすい。アドインで欲しいのは配る成果物だけ、と利用者に確かめた。

変換は、取り込みとリボンの埋め込みを済ませた一時コピーを Excel で開き、`Workbook.SaveAs(<path>, 55)`（`xlOpenXMLAddIn`）で保存する。ロックはその後に `.xlam` に掛ける。ロックの確認が、配る `.xlam` そのものを読むようにするため。

## 確かめた事実

- VBA の標準モジュールと `customUI/customUI14.xml` を持つ `.xlsm` を `SaveAs` で `.xlam` にすると、zip の中に `xl/vbaProject.bin` と `customUI/customUI14.xml` が残った。
- その `.xlam` に `Lock-VbaProject.ps1` を当てると、`.xlsm` と同じ手順でロックでき、開き直すと `VBProject.Protection` が `1` だった。
- ロックした `.xlam` を `Workbooks.Open` で開くと、`IsAddin` が `true`、`FileFormat` が `55` で、マクロ（`AddNums`）が動いた。
- `Test-RibbonLayout.ps1` をロックした `.xlam` に当てると、リボンが XML と合い、終了コードが 0 だった。
- 変換、`.xlsm` と `.xlam` のロック、両方のリボンの照合を、手順ごとに前面のウィンドウを 50 ms ごとに見ると、どの手順でも Excel は前面を取らなかった。
- 一方、`Test-XlWright.ps1 -Full -Case input-rejected,publish` の 1 回目は、`publish` の項目で 1059 回のうち 16 回、Excel が前面にあった。同じ項目を流し直すと 0 回で合格した。この 16 回は、試験の間に人が Excel のウィンドウに触れたためだった（人からの報告）。
- `.xlam` の中身を持たないファイル（`.xlsm` を名前だけ `.xlam` に写したもの）は、`Workbooks.Open` が `COMException` で失敗した。

## 未確認

- `Invoke-RibbonButton.ps1` で `.xlam` のリボンのボタンを押せるか。
- アドインとして導入した（`AddIns` に足した）時の挙動。導入は Excel の設定を変えるので、人が行う。
- ブックのイベント（`Workbook_Open`、`Workbook_AddinInstall`）が `.xlam` で動くか。
