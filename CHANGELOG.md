# Changelog

利用者に見える変更を、版ごとに書きます。形式は [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) に、版の付け方は [Semantic Versioning](https://semver.org/spec/v2.0.0.html) に従います。

## [Unreleased]

### Added

- `XlWright.psm1` に `Import-ExcelVba`、`Test-ExcelVbaCompile`、`Find-ExcelCellError` を足した。`Import-Vba.ps1`、`Test-VbaCompile.ps1`、`Find-CellError.ps1` と同じ処理を、`Invoke-ExcelScript.ps1` に渡す ScriptBlock の中で呼べる。取り込み、コンパイル、マクロの試験、エラー値のセルの検査を、1 回の Excel の起動で済ませられる。スクリプトは呼ぶたびに Excel を起動して閉じ、測った端末では 1 回に約 4 秒かかっていた。3 つのスクリプトの引数、出力、終了コードは変わらない。

### Changed

- `Publish-Workbook.ps1` が、配布用のブックから個人の情報（作成者、最終の編集者、管理者、会社、コメントの作成者、保存したフォルダのパス）を消す。出力の後に読み戻し、残っていれば失敗する。作業ブックは変えない。消さない時は `-KeepPersonalInformation` を付ける。今までは、作業ブックの情報がそのまま配布用のブックに残っていた。

### Fixed

- `Invoke-ExcelMacro` に、PowerShell の PSObject で包まれた値（`Join-Path` や `Get-Content` が返した文字列など）を `-ArgumentList` で渡すと、`DISP_E_TYPEMISMATCH`（`0x80020005`）で失敗していた。同じ値を `$xl.Run` に渡すと通っていた。中の値を渡すようにした。報告のあったアドインそのものでは確かめていない。
- Excel のオプション「アドインのユーザー インターフェイス エラーを表示する」がオンの時に、誤りのあるリボン XML のブックで `Test-RibbonLayout.ps1` と `Invoke-RibbonButton.ps1` が `WindowState` の `0x800AC472` で失敗していた。ブックを閉じる処理も拒まれ、Excel は強制的に止められていた。Excel が出すダイアログを閉じ、その文言（誤りの行と説明）を `Error` に入れて止まるようにした。
- `Save-WorkbookScreenshot.ps1` で画面より大きい `-Width` / `-Height` を指定すると、画像は黙って縮んでいた。縮んだ時は `Warnings` に出すようにした。文書の「画像も同じ大きさになる」も直した。
- `Test-Prerequisite.ps1 -Vba` が、sandbox などで `AccessVBOM` のレジストリを読めなかった時にも「値が 1 でない」と報告し、人に設定の変更を頼むよう案内していた。読めなかった時は別の問題として、設定ではなく権限の引き上げを求めるようにした。

### Added

- `Publish-Workbook.ps1` が、VBA をロックしたアドイン（`.xlam`）を出力できる。`-FileName` の拡張子を `.xlam` にする（`-FileName tool.xlam`）。作業ブックの形式は変わらない。
- `.xlsb` のブックを、`.xlsm` と同じように作り、編集し、配布できる。`New-ExcelWorkbook` と `-Create` で `.xlsb` を作れ、`Publish-Workbook.ps1` は `.xlsb` の作業ブックを受け付け、`.xlsb` でも出力できる。

### Changed

- **Breaking:** 扱うブックを `.xlsx`、`.xlsm`、`.xlsb`、`.xlam` に限った。ほかの形式（`.xltm`、`.xls` など）は、どのスクリプトもブックを開く前に拒否する。`Test-Prerequisite.ps1 -Path` も、その形式を問題として返す。今までは、Excel が開けるブックならそのまま開いていた。対象のブックを、Excel で `.xlsm` か `.xlsx` に保存し直してから渡す。

### Added

- Codex 用の manifest（`plugin.json`）と marketplace（`.agents/plugins/marketplace.json`）。ChatGPT の app の Codex タブで、この repo を marketplace として足して入れられる。

### Changed

- スクリプトの実行が実行ポリシー、sandbox、権限で拒否された時に、Excel を使わない手段で代わりをせず、権限の引き上げを求めるか止まるよう、スキルに書いた。
- `SKILL.md` のコマンド例で、スキルのフォルダを `${CLAUDE_SKILL_DIR}` ではなく `<skill>`（`SKILL.md` があるフォルダ）と書くようにした。`${CLAUDE_SKILL_DIR}` を展開しないエージェント（Codex など）でも、スクリプトのパスを探さずに決められるようにした。

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
