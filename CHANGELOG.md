# Changelog

利用者に見える変更を、版ごとに書きます。形式は [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) に、版の付け方は [Semantic Versioning](https://semver.org/spec/v2.0.0.html) に従います。

## [Unreleased]

### Added

- `XlWright.psm1` に `Import-ExcelVba`、`Test-ExcelVbaCompile`、`Find-ExcelCellError` を足した。`Import-Vba.ps1`、`Test-VbaCompile.ps1`、`Find-CellError.ps1` と同じ処理を、`Invoke-ExcelScript.ps1` に渡す ScriptBlock の中で呼べる。取り込み、コンパイル、マクロの試験、エラー値のセルの検査を、1 回の Excel の起動で済ませられる。スクリプトは呼ぶたびに Excel を起動して閉じ、測った端末では 1 回に約 4 秒かかっていた。3 つのスクリプトの引数、出力、終了コードは変わらない。
- 動作確認の試験（`tests/Test-XlWright.ps1`）に項目を足した。既定の項目は、ブックを開いてもマクロと `Workbook_Open` が動かないことを確かめる。`New-ExcelWorkbook` と `-Create` が既にあるブックを上書きしないことも確かめる。`Publish-Workbook.ps1` が、同名の出力、作業ブックと同じ場所の出力、`.xlsm` と `.xlsb` 以外の作業ブックを拒否することも確かめる。`-Full` の項目は、`-Force` で出力が置き換わることを確かめる。

### Changed

- `SKILL.md` に、スクリプトが受け付ける形を書いた。`Import-Vba.ps1` が受け付けるのは UTF-8（BOM なし）・LF のファイルだけで、`Publish-Workbook.ps1` が読むのはブックの隣の `vba/` と `ribbon/customUI14.xml` だけ。利用する repo の規則がこれと合わない時は、迂回せずにユーザーに伝える。今までは、repo の規則を優先するとだけ書いていた。
- `SKILL.md` に、references にしか無かった安全の規則を移した。Excel のオプションの変更とアドインの導入は、ユーザーに頼む。接続文字列はユーザーに確かめてから読む。書き出したコードは commit の前に確かめる。画面の画像は、ブックにも Git にも入れない。
- UserForm の実行時の確認は、一時フォルダに取った写しで、`Invoke-ExcelMacro` を使って行うと書いた（`references/userform.md`）。今までは `$xl.Run` で呼ぶと書いていて、フォームがダイアログを出すと上限まで止まるおそれがあった。
- `Lock-VbaProject.ps1` は、`Publish-Workbook.ps1` が作る配布の写し（名前が `.excel-publish-` で始まるファイル）だけを受け付け、ほかのブックはブックを開く前に拒否する。ロックのパスワードは捨てるので、作業ブックに直接当てるとロックを外せなくなっていた。配布は今までどおり `Publish-Workbook.ps1` で行う。
- リボンの文書（`references/ribbon.md`）に、実機で確かめた事実を書いた。`Invoke-RibbonButton.ps1` は toggleButton を押せるが、menu、menu の中の項目、dropDown は押せない。リボンの表示が「タブのみを表示する」の時、`Test-RibbonLayout.ps1` の照合は通るが、`-ScreenshotDir` の画像にリボンが写らない。「全画面表示モード」の時は、どちらのスクリプトもタブが見つからずに止まる。どちらの時も、ユーザーにリボンの表示を確かめてもらう。
- リボンの文書（`references/ribbon.md`）に、`.xlam` を閉じた後の振る舞いを書いた。閉じた時に前面になかったブックのウィンドウにはタブが残り、そのウィンドウを前面にすると消える。`.xlam` のマクロへの `Application.OnTime` は予約した時刻に動かないことがあり、`OnTime` で閉じ直すアドインは開いたままになる。タブが消えたかは、アドインが閉じたかを確かめ、確かめるウィンドウを前面にしてから見る。
- 新しく作るツールの置き場を決めた。repo の `src/excel/<ツール名>/` に、ブック `workbook.xlsm` と `vba/`、`ribbon/` を置き、配布用のブックは repo 直下の `dist/` に `<ツール名>.xlsm` として出力する。既にあるブックは、頼まれない限り移さない。利用する repo に置き場の規則があれば、そちらを優先する。スクリプトの引数と振る舞いは変わらない。今までは、ブックのフォルダの中の形だけを決め、そのフォルダを repo のどこに置くかは決めていなかった。
- UserForm の文書（`references/userform.md`）に、足したコントロールのフォントの扱いを書いた。足したコントロールは足した時のフォームの `Font` を受けるので、フォーム全体のフォントを変える時はフォームの `Font` も変える。コントロールを足した後は、一時の VBA で全コントロールのフォントを読み、フォームの `Font` と違うものがあればユーザーに確かめる。今までの文書には、足したコントロールのフォントを他のコントロールと比べる手順が無かった。
- 元にしたスキルから写した記述のうち、この repo で確かめていないものに、そう書いた。openpyxl で保存した時のリボン、`0x80080005` での起動の失敗、LF の `.cls` を VBE へ直接渡した時の扱い、日本語のモジュール名とパスが当たる。PowerShell 7 での振る舞いは確かめていないので、書くのをやめた。
- sandbox で Excel を起動できなかった例を、Codex の sandbox で確かめた結果に替えた。スクリプトの実行が `AuthorizationManager` の検査で止まり、Excel の起動も `0x80070005` で拒否された。
- ThisWorkbook に置いた関数は、`$xl.Run` だけでなく `Invoke-ExcelMacro` でも `null` が返ると書いた。値を読む時は、今までどおり `InvokeMember` で Workbook のメンバーとして呼ぶ。モジュール名の 31 文字の上限は、32 文字の名前が拒否されることを確かめた。
- `Publish-Workbook.ps1` が、配布用のブックから個人の情報（作成者、最終の編集者、管理者、会社、コメントの作成者、保存したフォルダのパス）を消す。出力の後に読み戻し、残っていれば失敗する。作業ブックは変えない。消さない時は `-KeepPersonalInformation` を付ける。今までは、作業ブックの情報がそのまま配布用のブックに残っていた。

### Removed

- `references/com.md` の「別ブックからシートをコピーするとき」を消した。元にしたスキルの repo の作業から写したもので、Excel を土台にしたツールの開発には当たらず、この repo では確かめていなかった。複数のブックを扱う時は、今までどおり「モジュールを直接使う」に従う。
- 同じ理由で、UserForm の文書（`references/userform.md`）から、元にしたツールの作りに寄った記述を消した。消したのは、旧コントロールを隠して固定レイアウトを作り直す処理の扱い、シートの表示状態とチェック欄の連動と `Caption` をシート名のキーに使う例、フォームだけを変えた時に `xl/vbaProject.bin` だけを元のブックの写しに移す手順。

### Fixed

- `SKILL.md` は、`Invoke-ExcelScript.ps1` に渡す ScriptBlock が返した複数セルの `Value2` が行の配列になると書いていた。そうなるのは hashtable に入れた時だけで、そのまま出力すると 1 列に平らになる（`A1:B2` が `[1,2,3,4]` になった）。文書を直した。
- 試験の `-Case` に無い名前を渡すか、`-Full` の項目を `-Full` なしで渡すと、何も流さずに合格（終了コード 0）になっていた。何も流さずに終了コード 2 で止まるようにした。
- 試験の前提の確認が、`AccessVBOM` のキーを読めなかった時（sandbox など）にも、設定を変えるようユーザーに頼むと案内していた。`Test-Prerequisite.ps1` と同じく、読めなかった時は権限の引き上げを求める案内にした。
- 配布の試験は、メモの作成者が消えたかを、Excel のユーザー名ではなくログイン名と比べていた。2 つが違う端末では、作成者が残っていても合格になった。Excel のユーザー名と比べ、作業ブックの側で作成者が入っていることも確かめるようにした。
- `Invoke-ExcelMacro` に、PowerShell の PSObject で包まれた値（`Join-Path` や `Get-Content` が返した文字列など）を `-ArgumentList` で渡すと、`DISP_E_TYPEMISMATCH`（`0x80020005`）で失敗していた。同じ値を `$xl.Run` に渡すと通っていた。中の値を渡すようにした。配列で渡した引数（VBA では Variant の配列になる）の要素も同じように扱う。報告のあったアドインそのものでは確かめていない。
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
