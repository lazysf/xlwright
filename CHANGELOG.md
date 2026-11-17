# Changelog

利用者に見える変更を、版ごとに書きます。形式は [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) に、版の付け方は [Semantic Versioning](https://semver.org/spec/v2.0.0.html) に従います。

## [Unreleased]

### Added

- VBA プロジェクトをコンパイルし、コンパイルエラーの文言と場所を返す `Test-VbaCompile.ps1`。
- マクロを動かし、実行中に出たダイアログ（実行時エラー、MsgBox、コンパイルエラー）の文言を返す `Invoke-ExcelMacro`。
- ブックの `Test_` で始まる Sub を VBA のテストとして順に実行し、合否とエラーの文言を返す `Invoke-VbaTest.ps1`。
- 正本の VBA とブックの中の VBA を比べ、差のあるコンポーネントを返す `Compare-Vba.ps1`。
- `Import-Vba.ps1` の `-Prune`。付けた時だけ、正本に無い標準モジュール、クラス、UserForm をブックから消し、正本に無い文書モジュールのコードを空にする。
- `Publish-Workbook.ps1` の `-ExcludeTests`。付けた時だけ、名前が `Test_` で始まる標準モジュールを配布用のブックから消す。作業ブックには残る。
- 作業の前提（Excel のプロセス、`AccessVBOM`、WMI、ブックがほかで開かれていないか、Git の状態）を、Excel を起動せずに調べて返す `Test-Prerequisite.ps1`。
- ブックのシート、定義名、テーブル、外部リンク、データ接続、VBA のコンポーネント、customUI を 1 回で返す `Get-WorkbookSummary.ps1`。
- セルの範囲を行の配列として 1 回で読み書きする `Read-ExcelRange` と `Write-ExcelRange`。文字列は既定で文字のまま書き、`-Formula` を付けると数式として書く。エラー値は名前（`#DIV/0!` など）で返す。
- 値がエラー（`#REF!`、`#NAME?` など）になったセルを、シート、番地、エラーの名前、数式とともに返す `Find-CellError.ps1`。

## [0.1.0] - 2026-11-13

### Added

- Excel のブックを Windows 版 Excel と PowerShell の COM で確認・編集するスキル `xlwright` を、Claude Code のプラグインとして配布する。
- シートの値・数式・書式の確認と編集。
- VBA（標準・クラス・文書モジュール）と UserForm の書き出し・編集・取り込み。
- カスタムリボンの埋め込みと、見え方の照合、ボタンを押して動作を確かめること。
- VBA プロジェクトをロックした配布用のブックの出力。
- 導入先の Excel で動くかを確かめる試験 `tests/Test-XlWright.ps1`。
