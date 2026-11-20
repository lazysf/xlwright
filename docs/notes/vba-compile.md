# VBA のコンパイルを COM から起こし、エラーを読む

- Date: 2026-11-16
- 出典: 一時フォルダの実験用スクリプトを 3 回流した実測。Microsoft 365 Excel 64-bit 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818、画面の拡大率 100%。Claude のウィンドウを前面にして、人は何も触らない状態で測った。

VBA のコンパイルには、COM の直接の API が無い。VBE のメニューの「VBAProject のコンパイル」を `CommandBars` から実行すると、Windows PowerShell 5.1 と `Use-ExcelWorkbook`（非表示の Excel）でもコンパイルでき、エラーの文言と場所を読めた。

実装は `plugins/xlwright/skills/xlwright/scripts/XlWright.psm1` の `Test-ExcelVbaCompile` を正本とする。`Test-VbaCompile.ps1` はこれを呼ぶ。使い方は `plugins/xlwright/skills/xlwright/references/vba.md` にある。

同じ方法を .NET から確かめた記録が、`mcp` ブランチの `docs/notes/test-copy.md` と `docs/notes/dialog-probe.md` にある。

## 実験の手順

場合ごとに新しいブックと別の Excel を使った。

1. `New-ExcelWorkbook` で `.xlsm` を作り、`Use-ExcelWorkbook` で開く。
2. 標準モジュール `Probe` を足し、`CodeModule.AddFromString` で 5 行のコードを入れる。
3. `$xl.VBE.ActiveVBProject` にそのブックの `VBProject` を入れる。
4. `$xl.VBE.CommandBars.FindControl(1, 578)` の `Execute()` を呼ぶ。
5. 5 秒の間、100 ms ごとに、その Excel のプロセスが持つ、表示されたトップレベルのウィンドウ（クラス `#32770`）を `EnumWindows` で探す。
6. 見つけたら、題名と子ウィンドウの文字を `GetWindowText` で読み、`ActiveCodePane.GetSelection` を読む。`WM_CLOSE` を `PostMessage` で送り、`VBE.MainWindow.Visible` を偽にする。

場合は 3 つ。正しいコード、`Option Explicit` での宣言の無い変数（4 行目）、構文エラー（4 行目の `a = = 1`）。

## 結果

3 回とも同じだった。

| 見たもの | 正しいコード | 宣言の無い変数 | 構文エラー |
| --- | --- | --- | --- |
| `Execute()` が戻るまで | 3 ms | 0 ms | 0 ms |
| ダイアログ | 5 秒の間、出ない | 109〜220 ms で見つかった | 103〜118 ms で見つかった |
| コントロールの `Enabled` | 真から偽になった | 真のまま | 真のまま |
| `VBE.MainWindow.Visible` | 偽のまま | 真になった | 真になった |
| `ActiveCodePane.GetSelection` | 読んでいない | `Probe` の 4 行目、5 列目から | `Probe` の 4 行目、1 列目から |

- コントロールの `Caption` は「VBAProject のコンパイル(&L)」だった。
- ダイアログの題名は「Microsoft Visual Basic for Applications」。子ウィンドウは、`Button` の「OK」と「ヘルプ」、`Static` の本文だった。本文は「コンパイル エラー:」、空行、説明（「変数が定義されていません。」「構文エラー」）の順だった。
- `Execute()` はダイアログが出ていても戻る。ダイアログが出ている間も、COM の呼び出し（`Enabled`、`MainWindow.Visible`、`GetSelection`）は通った。別の runspace から呼ぶ必要は無かった。
- `WM_CLOSE` でダイアログは閉じた（測った 1 回では 59 ms と 72 ms）。`MainWindow.Visible` を偽にすると、VBE は隠れた。
- ダイアログを見つけた時（正しいコードでは 5 秒の後）の前面のウィンドウは、9 回とも Claude のままだった。見たのはその 1 点だけで、間は見ていない。
- 各回の後に Excel のプロセスは残らず、`Use-ExcelWorkbook` の警告も無かった。
- 構文エラーの場合の `Use-ExcelWorkbook` の全体は、測った 2 回で 11.6 秒と 3.5 秒だった。

## スクリプトで確かめたこと

`Test-VbaCompile.ps1` を、`Import-Vba.ps1` で取り込んで保存したブックに当てた（2026-11-16）。

- マクロを無効にして開いた（`AutomationSecurity` が 3）既存のブックでも、コンパイルは起き、エラーのダイアログが出た。
- VBA の無い `.xlsx` と、モジュールを足していない `.xlsm` は、コンパイルの済んだ状態になった。
- `GetSelection` の行は、コードウィンドウの行だった。`.bas` の 6 行目（先頭が `Attribute VB_Name` の行）のエラーは、5 行目と読めた。
- スクリプトの全体は、1 回で 4〜12 秒だった。エラーの有無によらず、12 秒かかる回があった。
- `plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` の `vba-compile-error` は、前面のウィンドウを 50 ms ごとに見る。宣言の無い変数と構文エラーの検査の間、Excel は一度も前面に来なかった。

合格をコントロールの `Enabled` で見分けるのは、ダイアログが出ないことだけでは、コンパイルが起きなかった場合と区別できないため。どちらも確かめられない時は、合格にせず失敗にしている。

`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` の `vba-compile` と `vba-compile-error` が、これを確かめる。ダイアログを見つけられなくした写しは `vba-compile-error` で、済んだ状態を見ないようにした写しは `vba-compile` で落ちた。

## mcp ブランチの記録との違い

`mcp` ブランチの記録では、構文エラーの時に `ActiveCodePane` が無いことが 3 回中 2 回あった。今回は 3 回とも読めた。違いの理由は確かめていない。場所は読めないことがあるものとして扱う。

## 未確認

- エラーが複数ある時。コンパイルは最初のエラーで止まると見込むが、確かめていない。
- モジュールが複数ある時、文書モジュールや UserForm にエラーがある時の `ActiveCodePane`。
- VBA プロジェクトがロックされたブック。スクリプトは `VBProject.Protection` を見て、コンパイルせずに失敗にするが、その動きは確かめていない。
- VBE のウィンドウが先に表示されている時。
- 構文エラーの時に `Use-ExcelWorkbook` の全体が 10 秒を超えた理由。`Quit` で Excel が終わらず、`Stop-ExcelSession` が自分の Excel を止めたと推測している。
- 大きなプロジェクトで、ダイアログが出るまでの時間が 5 秒を超えるか。
- 32-bit の Excel、日本語以外の表示言語（ダイアログの文言と `Caption` が変わる）。
- コンパイルエラーのダイアログが 3 秒で閉じない時に、Excel を止める分岐。試験ではダイアログが閉じるので、この分岐は流れていない。照合して止める `XlWright.ExcelProcess.StopIfSame` そのものは、`stop-own-process` が確かめる。
