# VBA の取り込み、コンパイル、マクロを 1 つの Excel で行う

- Date: 2026-11-20
- 出典: 一時フォルダの実験用スクリプトを場合ごとに 1 回流した実測。Microsoft 365 Excel 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818（64-bit）。人は何も触らない状態で測った。

スクリプトの呼び出しは、1 回ごとに Excel を起動して閉じる時間が約 4 秒かかる（[dev-cycle-timing.md](./dev-cycle-timing.md)）。そこで、`Import-Vba.ps1`、`Test-VbaCompile.ps1`、`Find-CellError.ps1` の処理を `XlWright.psm1` の関数（`Import-ExcelVba`、`Test-ExcelVbaCompile`、`Find-ExcelCellError`）に移し、1 つの ScriptBlock から続けて呼べるようにした。今までの `Test-VbaCompile.ps1` は、マクロを無効にして開いた Excel でしか使っていなかった。移す前に、取り込みの直後とマクロを有効にした Excel でもコンパイルを確かめられるかを試した。

使い方は `plugins/xlwright/skills/xlwright/SKILL.md` の「まとめて確かめる」にある。

## 実験の手順

場合ごとに、`New-ExcelWorkbook` で新しい `.xlsm` を作った。`Use-ExcelWorkbook -EnableMacros` で開き、1 つの ScriptBlock の中で次を行った。

- 取り込みは、`tests/data/` の `.bas` を CP932・CRLF の一時ファイルにして `VBComponents.Import` に渡した。
- コンパイルは、`Test-VbaCompile.ps1` と同じ手順（VBE の「VBAProject のコンパイル」を実行し、ダイアログを探して閉じる）で行った。
- マクロは `Invoke-ExcelMacro` で動かした。

| 場合 | ScriptBlock の中の順 |
| --- | --- |
| 正しいコード | `vba/standard/Util.bas` と `vba-macro/standard/Macros.bas` を取り込む → コンパイル → `Macros.AddThree(1, 2, 3)` → `Util.AddNums(2, 3)` → もう一度コンパイル → 保存 |
| 宣言の無い変数 | `vba-broken/undeclared/` の `Broken.bas` と `Good.bas` を取り込む → コンパイル → `Good.Seven` → もう一度コンパイル |
| 構文エラー | `vba-broken/syntax/` の 2 つで、上と同じ順 |
| マクロを先に動かす | `vba-broken/undeclared/` の 2 つを取り込む → `Good.Seven` → コンパイル |

正しいコードの場合は、別の新しいブックに同じ 2 つのファイルを `Import-Vba.ps1` で取り込んだ。2 つのブックを `Export-Vba.ps1` で書き出し、`diff -r` で比べた。

## 結果

| 場合 | コンパイル | マクロ | 2 回目のコンパイル |
| --- | --- | --- | --- |
| 正しいコード | 通った | 6 と 5 が返った | コンパイルの済んだ状態（コントロールが無効）で、通った |
| 宣言の無い変数 | 「コンパイル エラー: 変数が定義されていません。」、`Broken` の 5 行目 | `Good.Seven` が 7 を返した | 1 回目と同じエラー |
| 構文エラー | 「コンパイル エラー: 構文エラー」、`Broken` の 5 行目 | `Good.Seven` が 7 を返した | 1 回目と同じエラー |
| マクロを先に動かす | 「コンパイル エラー: 変数が定義されていません。」、`Broken` の 5 行目 | 先に動かした `Good.Seven` が 7 を返した | - |

- コンパイルエラーのダイアログは、どの場合も `WM_CLOSE` で閉じた。
- エラーのあるモジュールが残っていても、ブックは警告なしで閉じた。
- 正しいコードの場合、`Import-Vba.ps1` で取り込んだブックと書き出しを比べると、差は無かった。比べたのは `Util.bas`、`Macros.bas`、`references.json`。
- 1 つの ScriptBlock で済ませた `Use-ExcelWorkbook` の全体は、4.0〜6.4 秒だった。正しいコードの場合とマクロを先に動かした場合では、ScriptBlock の終わりから戻るまでが約 5.2 秒で、戻った直後に Excel のプロセスが 1 つ残っていた（その後の確認では 0）。`Quit` の後 5 秒の待ちを過ぎ、`Stop-ExcelSession` が止めたと見ている。区間は測っていない。残りの 2 つの場合は、戻るまでが約 2.6 秒で、直後に Excel は残っていなかった。

## 関数に移した後に確かめたこと

`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` に、次の確認を足した。

- `one-session`: ブックの隣の `vba/` を、`Invoke-ExcelScript.ps1 -EnableMacros -Save` の 1 回の呼び出しで、取り込み、コンパイル、マクロの実行、エラー値のセルの検査まで行う。保存したブックを開き直し、取り込んだコンポーネントと、同じ ScriptBlock で書いたセルが残っていることも見る。
- `vba-compile-error`（`-Full`）: 宣言の無い変数と構文エラーのそれぞれで、`Test-ExcelVbaCompile` がエラーと場所を返すこと、そのあと同じ Excel で `Good.Seven` が 7 を返すことを見る。

この確認が壊れたものを見つけられるかは、関数をわざと壊した写しで確かめた（2026-11-20）。

- 取り込みで `VBComponents.Import` を呼ばない写しは、`one-session` で落ちた。
- エラー値のセルを数えない写しは、`one-session` で落ちた。
- コンパイルせずに「通った」と返す写しは、`vba-compile-error` で落ちた。`one-session` は正しいコードだけを使うので、この写しでは落ちない。

## 確かめていないこと

- コンパイルエラーの後に、エラーのあるモジュールのマクロを `Invoke-ExcelMacro` で動かした時の振る舞い。
- UserForm、文書モジュール、クラスにエラーがある時。
- `Quit` の後に Excel が終わらなかった理由。この回の Excel の診断ログは残っておらず、[excel-quit.md](./excel-quit.md) で見たライセンスの更新に当たるかは分からない。
