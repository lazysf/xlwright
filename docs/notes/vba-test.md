# VBA のテストを列挙して実行する

- Date: 2026-11-16
- 出典: 一時フォルダの実験用スクリプトを 3 回流した実測。Microsoft 365 Excel 64-bit 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818。Claude のウィンドウを前面にして、人は何も触らない状態で測った。

ブックの標準モジュールのコードを読むと、テストの `Sub` を列挙できた。`Invoke-ExcelMacro` で 1 つずつ呼ぶと、失敗したテストの `Err.Raise` の説明を、実行時エラーのダイアログから読めた。1 つの Excel で、失敗するテストを続けて実行できた。

実装は `plugins/xlwright/skills/xlwright/scripts/Invoke-VbaTest.ps1` を正本とする。使い方は `plugins/xlwright/skills/xlwright/references/vba.md` の「テストの実行」にある。ダイアログの扱いは [macro-dialog.md](./macro-dialog.md) にある。

## 実験の手順

1. 2 つの標準モジュールを `Import-Vba.ps1` で取り込んだ `.xlsm` を、`Invoke-ExcelScript.ps1 -EnableMacros` で開く。
2. 標準モジュール（`VBComponent.Type` が 1）ごとに、`CodeModule.Lines` で全文を読む。行の継続（`_`）をつなぎ、行の先頭の `Sub` の宣言を正規表現で探す。比べるために、`CodeModule.ProcOfLine` でもプロシージャを列挙する。
3. `Invoke-ExcelMacro` で、通る `Sub`、`Err.Raise` で失敗する `Sub`、0 で割る `Sub` を順に呼び、時間と例外の文言を記録する。0 で割る `Sub` は、最後に 5 回続けて呼ぶ。
4. モジュール変数を増やす `Sub` と、その値を返す関数で、失敗の前後の値を読む。

## 結果

3 回とも同じだった。時間だけが回ごとに違った。

### 列挙

- 正規表現は、`Public Sub`、`Public` を省いた `Sub`、`Public Static Sub`、宣言を行の継続で分けた `Sub`、日本語の名前の `Sub` を拾った。`Private Sub`、引数（`Optional` を含む）のある `Sub`、`Function`、コメントの中の宣言は拾わなかった。
- `ProcOfLine` は、`Private` も `Function` も同じ種別（0）で返し、引数の有無も分からなかった。列挙には使っていない。
- `VBComponents` は、2 つのモジュールをモジュール名の順に返した。取り込んだ順も同じだったので、どちらの順なのかは分からない。スクリプトは、モジュール名で並べ直す。

### 実行

| 呼んだもの | 結果 | 時間 |
| --- | --- | --- |
| 通る `Sub` | 例外なし | 14〜30 ms |
| `Err.Raise vbObjectError + 1, , "expected 3, actual 2"` | 文言は「実行時エラー '-2147221503 (80040001)': expected 3, actual 2」 | 77〜197 ms |
| `Err.Raise 513, , "plain 513"` | 文言は「実行時エラー '513': plain 513」 | 同上 |
| 説明が 2 行で、300 文字を超える `Err.Raise` | 説明は欠けずに読めた。改行は空白 1 つになった | 同上 |
| 0 で割る | 文言は「実行時エラー '11': 0 で除算しました。」 | 同上 |
| 0 で割る `Sub` を 5 回続けて | 5 回とも同じ文言で例外になり、その後の通る `Sub` は通った | 各 77〜189 ms |

- 日本語の説明（「期待は 3、実際は 2」）も、そのまま読めた。
- `Err.Raise` の `Source` は、ダイアログに出ない。
- `Private Sub` は、`<モジュール>.<Sub>` の形で呼ぶと実行された。`Private` を除くのは、列挙の側だけである。
- `Optional` の引数だけの `Sub` は、引数なしで実行された。必須の引数のある `Sub` は、`DISP_E_PARAMNOTOPTIONAL` で失敗した。
- 2 つのモジュールに同じ名前の `Sub` がある時、モジュール名を付けない呼び出しは「マクロを実行できません」で失敗した。モジュール名を付けると、それぞれ実行された。
- 実行時エラーを「終了」で閉じると、モジュール変数は初期値に戻った。2 回増やして 2 だった値が、失敗の後に 0 になった。
- 実験のスクリプトの全体（マクロの呼び出しは 27 回、うち失敗は 13 回）は、測った 1 回で 6 秒だった。Excel のプロセスは、3 回とも残らなかった。

## Invoke-VbaTest.ps1 の形にした理由

- **テストをそのまま `Invoke-ExcelMacro` で呼ぶ。** `mcp` ブランチは、`On Error` で包む関数を実行ごとにブックへ作り、`Err.Number` と `Err.Source` を受け取る（`mcp` ブランチの `docs/api.md` の `test_run`）。こちらはブックに何も足さない。代わりに、失敗するテストごとに実行時エラーのダイアログが画面に短い間出る。エラーの番号と説明は、ダイアログの文言として 1 つの文字列で返る。
- **アサーションのモジュールを同梱しない。** テストは `Err.Raise` で失敗を知らせる。そのため、アサーションの失敗とほかの実行時エラーを区別せず、どちらも `failed` にする。
- **テストはブックの中に置く。** 写しへ取り込む形にしない。配布ブックにテストのモジュールが残る点は、`references/vba.md` に書いた。
- **`Private Sub` を名前で指定しても実行しない。** `-Name` は、列挙したテストの中からだけ選ぶ。`Application.Run` は `Private Sub` も実行するため。
- **コンパイルエラーのテストは `failed` にして、先へ進む。** `Invoke-ExcelMacro` が VBE をリセットし、Excel が使えるまま戻るため（[macro-dialog.md](./macro-dialog.md)）。
- **Excel を止めた後のテストは `not_run` にする。** 閉じられないダイアログと上限では、`Invoke-ExcelMacro` と見張りがその Excel を止める。残りのテストは実行できない。

`tests/Test-XlWright.ps1` の `vba-test` と `vba-test-fail` が、これを確かめる（2026-11-16）。実験に加えて、次を確かめた。

- テストの実行で、ブックのファイルは変わらなかった。
- OK だけの MsgBox を出すテストは `passed` になり、文言が `Warnings` に入った。
- 宣言の無い変数のあるモジュールのテストは `failed` になり、`Message` にコンパイルエラーの文言が入った。後のテストは実行された。
- Yes/No の MsgBox を出すテストは `failed` になり、後のテストは `not_run`、`Error` に Excel を止めたことが入った。
- 終わらないテストは、`-TimeoutSeconds 20` で `failed` になり、後のテストは `not_run`、`Error` は `Timed out after 20 s` で始まった。
- `vba-test-fail` は前面のウィンドウを 50 ms ごとに見る。Excel は一度も前面に来なかった。
- `Private Sub` も拾うようにした写しと、Excel が止まった後も続けるようにした写しは、`vba-test` で落ちた。`Invoke-ExcelMacro` の例外から種別と文言を外した写しは、`vba-test-fail` で落ちた。

## 未確認

- VBA プロジェクトがロックされたブック。スクリプトは `VBProject.Protection` を見て失敗にするが、その動きは確かめていない。
- 取り込んだ順とモジュール名の順が違う時に、モジュール名の順に実行されること。試験のデータも、2 つの順が同じである。
- テストが `On Error` で自分のエラーを握りつぶした時。`passed` になる見込み。
- テストの数が多い時（100 個以上）の時間。
- 「エラー発生時に中断」を VBE のオプションで選んでいる時。
- 32-bit の Excel、日本語以外の表示言語。
