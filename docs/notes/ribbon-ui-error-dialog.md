# 拒否されたリボン XML と「アドインのユーザー インターフェイス エラーを表示する」

- Date: 2026-11-20
- 出典: 一時フォルダの新しい `.xlsm` で流した実測と、試験の `ribbon-layout`。Excel 16.0.20527.20196（Microsoft 365 64-bit）、Windows 11、Windows PowerShell 5.1.26100.9818、画面の拡大率 100%。

Excel のオプション「アドインのユーザー インターフェイス エラーを表示する」（「ファイル」→「オプション」→「詳細設定」→「全般」）がオンだと、拒否されたリボン XML を持つブックを画面に出して開いた時に、Excel がダイアログを出す。このダイアログが開いている間、COM の呼び出しは拒まれる。

## 実験の手順

1. 新しい `.xlsm` に、試験の `tests/data/ribbon/layout-rejected.xml`（`button` に未定義の属性 `sise` がある）を `Add-CustomUI.ps1` で埋め込む。
2. `Test-RibbonLayout.ps1 -FindTimeoutSeconds 5` を流す。
3. 同じブックを `Use-ExcelWorkbook -Visible -EnableMacros` で開き、3 秒待つ。その Excel のプロセスが持つ、見えているトップレベルのウィンドウを `EnumWindows` で列挙する。続けて `WindowState` を `xlMaximized` にする。

オプションの切り替えは、人が行った。手順を、オプションがオフの時とオンの時に 1 回ずつ流した。

## 結果

| オプション | `Test-RibbonLayout.ps1` | 開いた時のウィンドウ | `WindowState` の設定 |
| --- | --- | --- | --- |
| オフ | 終了コード 1。タブが `missing` | `XLMAIN` だけ | 通った |
| オン | 終了コード 2。`Error` は `WindowState` の設定で `0x800AC472` | `XLMAIN` と `NUIDialog` | `0x800AC472` |

- オンの時のダイアログは、クラスが `NUIDialog` のトップレベルのウィンドウだった。`XLMAIN` の子ではない。題は `<ブック名> のカスタム UI 実行時エラー`。
- UI Automation で読んだ本文には、XML のパス、行と列（`行: 7`、`列: 99`）、エラーコード `0x80004005`、誤りの説明（要素 `button` の属性 `sise` は DTD でもスキーマでも定義されていない）が入っていた。ボタンは `OK` と `すべて OK` だった。
- オンの時は、後始末の `Workbook.Close` が `RPC_E_CALL_REJECTED`（`0x80010001`）、`Quit` が `0x800AC472` で失敗した。Excel は数秒後には終わっていた。後始末の `Stop-ExcelSession`（`XlWright.psm1`）が、5 秒待った後に PID と起動時刻で止めたと見ている。止めた時刻は記録していない。
- オフの時は、試験の `ribbon-layout` の「rejected XML: exit 1, tab missing」が期待するとおりの結果だった。オンの時は、この項目が落ちる。

外の環境（Excel 32bit 16.0.20430）で、同じ XML で同じ `0x800AC472` が出たという報告がある（人からの報告。その環境のオプションの状態は報告に無い）。

続けて、オンのまま同じブックを開き、`WindowState` が `0x800AC472` で失敗した直後に `NUIDialog` を探した。1 つ見つかり、UI Automation の `WindowPattern.Close` で閉じると消えた。閉じる前後で前面のウィンドウは変わらなかった。閉じた後は `WindowState` の設定が通り、ブックは閉じて Excel も終わった（開いてから 7 秒）。

## 今の作り

- `Set-RibbonWindow`（`Ribbon.psm1`）は、`WindowState` の設定が `0x800AC472` で失敗した時だけ、その Excel のプロセスの `NUIDialog` を 3 秒まで探す。見つかれば文言を読んで閉じ、`Excel rejected the custom UI XML:` に文言を続けて投げる。見つからなければ、元の例外を投げる。ダイアログが無い時は、探す手間がかからない。
- 探すのは、見えている `NUIDialog` だけにしている。1 つずつ閉じ、閉じた後に窓が消えたことを確かめてから、次を探す。上限は 5 つ。
- 閉じられなかった時（`WindowPattern.Close` が失敗した時、または 3 秒たっても窓が残った時）も、読めた文言は捨てない。文言に `(the dialog could not be closed: <理由>)` を付けて投げる。その後の後始末の `Stop-ExcelSession` が、その Excel を PID と起動時刻で止める。
- `WindowState` の設定が通った時のために、もう 1 か所でも探す。`Test-RibbonLayout.ps1` は XML のタブが待っても揃わなかった時に、`Invoke-RibbonButton.ps1` はタブが見つからなかった時に、`NUIDialog` を待たずに 1 回だけ探す（`Assert-NoCustomUIError`）。ダイアログが `Set-RibbonWindow` より後に出た場合に備えたもので、そうした出方は確かめていない。
- 閉じた後も照合を続けずに止める。拒否された XML ではカスタムのタブが作られないので、照合しても全部 `missing` になるだけで、誤りの箇所はダイアログの文言にしか無いからだ。
- 閉じるのは、後始末で `Workbook.Close` と `Quit` を通すためだ。閉じないと、どちらも拒まれ、Excel を強制的に止めることになる。
- `Test-RibbonLayout.ps1` と `Invoke-RibbonButton.ps1` は、どちらも `Set-RibbonWindow` を通るので、同じ扱いになる。試験の `ribbon-layout` が両方を確かめる。どちらの結果を確かめるかは、オプションの状態で決まる。試験はオプションを変えない。

## 未確認

- 非表示のまま開いた時（`Save-WorkbookScreenshot.ps1` は、非表示のまま `WindowState` を設定する）にもダイアログが出るか。
- 誤りが複数ある XML での、ダイアログの数と本文。
- 拒否された XML のブックを開いた時に、`Workbooks.Open` がダイアログで止まることがあるか。今回は `Open` は戻った。
- ダイアログが `WindowState` の設定より後に出ることがあるか。
- 見えていない `NUIDialog` を Excel が持つことがあるか。
- `WindowPattern.Close` が失敗する、または閉じても窓が残る場合が実際にあるか。その時の扱いは、試験で確かめていない。
