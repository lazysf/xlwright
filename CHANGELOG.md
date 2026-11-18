# Changelog

利用者に見える変更を、版ごとに書きます。形式は [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) に、版の付け方は [Semantic Versioning](https://semver.org/spec/v2.0.0.html) に従います。

## [Unreleased]

### Added

- VBA プロジェクトをコンパイルし、コンパイルエラーの文言と場所を返す `Test-VbaCompile.ps1`。
- マクロを動かし、実行中に出たダイアログ（実行時エラー、MsgBox、コンパイルエラー）の文言を返す `Invoke-ExcelMacro`。
- `Export-Vba.ps1` が、VBA の参照設定を `references.json` に書き出す。`Import-Vba.ps1` は、`references.json` にあってブックに無い参照設定を警告する（参照設定は足さない）。
- 作業の前提（Excel のプロセス、`AccessVBOM`、WMI、ブックがほかで開かれていないか、Git の状態）を、Excel を起動せずに調べて返す `Test-Prerequisite.ps1`。
- ブックのシート、定義名、テーブル、外部リンク、データ接続、VBA のコンポーネント、customUI を 1 回で返す `Get-WorkbookSummary.ps1`。
- セルの範囲を行の配列として 1 回で読み書きする `Read-ExcelRange` と `Write-ExcelRange`。文字列は既定で文字のまま書き、`-Formula` を付けると数式として書く。エラー値は名前（`#DIV/0!` など）で返す。
- 値がエラー（`#REF!`、`#NAME?` など）になったセルを、シート、番地、エラーの名前、数式とともに返す `Find-CellError.ps1`。
- ブックを開いた Excel のウィンドウを、画面に出さずに PNG に保存する `Save-WorkbookScreenshot.ps1`。リボン、セルの書式、グラフ、ボタン、シートの見出しを、人が見るのと同じ形で確かめられる。`-UserForm` を付けると、`UserForm_Initialize` の後の UserForm を画像にする（`-EnableMacros` も要る）。

### Changed

- スキルの `description` を、Excel を土台にしたツール（マクロ、UserForm、カスタムリボンを持つブック）の開発に絞った。受け取ったブックの分析、集計の確認、データの整形では、スキルが選ばれにくくなる。プラグインの説明も同じ範囲にそろえた。
- `Use-ExcelWorkbook` と `New-ExcelWorkbook` が、`Quit` の後に Excel の終了を待つ時間を、10 秒から 5 秒にした。`Quit` で Excel が終わらない時（測った範囲では、ほとんどの回）は、呼び出しが約 5 秒早く戻る。

### Fixed

- `Invoke-RibbonButton.ps1` が、OK/キャンセル の MsgBox を × で閉じ、マクロの答えを決めていた。答えを選ぶダイアログは閉じずに、そのスクリプトが起動した Excel を止めてエラーにする。
- `Invoke-RibbonButton.ps1` と `Lock-VbaProject.ps1` が、Excel を PID だけで照合して止めていた。起動時刻も照合する。

## [0.1.0] - 2026-11-13

### Added

- Excel のブックを Windows 版 Excel と PowerShell の COM で確認・編集するスキル `xlwright` を、Claude Code のプラグインとして配布する。
- シートの値・数式・書式の確認と編集。
- VBA（標準・クラス・文書モジュール）と UserForm の書き出し・編集・取り込み。
- カスタムリボンの埋め込みと、見え方の照合、ボタンを押して動作を確かめること。
- VBA プロジェクトをロックした配布用のブックの出力。
- 導入先の Excel で動くかを確かめる試験 `tests/Test-XlWright.ps1`。
