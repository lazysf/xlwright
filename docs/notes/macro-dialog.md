# マクロの実行中に出たダイアログを見つけて閉じる

- Date: 2026-11-16
- 出典: 一時フォルダの実験用スクリプトを 3 回流した実測。Microsoft 365 Excel 64-bit 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818、画面の拡大率 100%。Claude のウィンドウを前面にして、人は何も触らない状態で測った。

`$xl.Run` は、マクロがダイアログを出している間は戻らない。非表示の Excel でも、別の runspace から Excel のプロセスのウィンドウを見張ると、ダイアログを見つけて文言を読めた。実行時エラーと OK だけの MsgBox は、閉じて `Run` を戻せた。Yes/No の MsgBox は、`WM_CLOSE` では閉じず、戻せなかった。コンパイルエラーは、閉じただけでは戻らず、閉じた後に VBE をリセットすると戻せた。

実装は `plugins/xlwright/skills/xlwright/scripts/XlWright.psm1` の `Invoke-ExcelMacro` を正本とする。使い方は `plugins/xlwright/skills/xlwright/SKILL.md` の「マクロを動かす」にある。

同じことを .NET から確かめた記録が、`mcp` ブランチの `docs/notes/dialog-probe.md` にある。

## 実験の手順

1. `Import-Vba.ps1` で標準モジュールを取り込んだ `.xlsm` を、`Use-ExcelWorkbook -EnableMacros` で開く。場合ごとに別の Excel を使った。
2. 別の runspace（`[PowerShell]::Create()`）で、100 ms ごとに、その Excel のプロセスが持つ、表示されたトップレベルのウィンドウ（クラス `#32770`）を `EnumWindows` で探す。
3. 元の runspace で `$xl.Run('<モジュール>.<関数>')` を呼ぶ。
4. 見張りは、見つけたダイアログの題名と、子ウィンドウのクラス、コントロールの ID（`GetDlgCtrlID`）、文字を読む。その後、`WM_CLOSE` を送るか、ID を指定したボタンに `BM_CLICK` を送る。3 秒たってもダイアログが残れば、その Excel を PID で止める。
5. `Run` が戻った後に、値を返すだけの別のマクロを `Run` で呼び、Excel が使える状態かを見る。

## 結果

| マクロ | ダイアログ | 閉じ方 | `Run` |
| --- | --- | --- | --- |
| 値を返すだけ | 出ない | なし | 4 ms で値が返った |
| `Application.Wait` で 3 秒待つ | 出ない | なし | 2.3 秒で値が返った |
| OK だけの MsgBox | 101〜111 ms で見つかった | `WM_CLOSE` で閉じた | 119〜139 ms で値が返った（3 回） |
| 0 で割る（実行時エラー） | 11〜116 ms で見つかった | `WM_CLOSE` では閉じない（1 回）。「終了」に `BM_CLICK` で閉じた（3 回） | 64〜129 ms で例外 `0x800A9C68` になった |
| 宣言の無い変数（コンパイルエラー） | 215〜218 ms で見つかった | `WM_CLOSE` でも「OK」に `BM_CLICK` でも閉じた（各 1 回） | 戻らない。上限の 60 秒で見張りのプロセスが Excel を止めた |
| Yes/No の MsgBox | 111 ms で見つかった | `WM_CLOSE` では閉じない（1 回） | 戻らない。見張りが Excel を止めた |

- ダイアログの出ないマクロを、ダイアログと見誤らなかった。
- 実行時エラーのダイアログは、題名が「Microsoft Visual Basic」。子ウィンドウは次のとおり。ボタンの文字は表示言語で変わるので、押すボタンは ID で選ぶ。

  | クラス | ID | 文字 |
  | --- | --- | --- |
  | `Static` | 4803 | 実行時エラー '11': 0 で除算しました。（番号と説明の間に空行） |
  | `Button` | 4802 | 継続(&C) |
  | `Button` | 4800 | 終了(&E) |
  | `Button` | 4801 | デバッグ(&D) |
  | `Button` | 4902 | ヘルプ(&H) |

- 「終了」で閉じた後は、VBE のウィンドウは表示されず、別のマクロは値を返した（3 回）。Excel はそのまま使えた。
- コンパイルエラーのダイアログは、題名が「Microsoft Visual Basic for Applications」。子ウィンドウは `Button` の「OK」（ID 2）と「ヘルプ」（ID 9）、`Static` の本文（ID 65535）。ダイアログと同時に VBE のウィンドウ（`wndclass_desked_gsk`）が表示された。
- コンパイルエラーのダイアログを閉じた 1 秒後、VBE のウィンドウの題名に「[中断]」が付いていた。VBE が中断モードに入り、`Run` はそのまま戻らなかった。
- MsgBox の子ウィンドウは、`Static`（ID 65535）の本文と、`Button` の「OK」（ID 2）、または「はい」（ID 6）と「いいえ」（ID 7）。
- ダイアログを見つけた時の前面のウィンドウは、どの回も Claude のままだった。見たのはその 1 点だけ。
- 見張りが Excel を止めた時、`Run` は `0x800706BE` の例外になり、`Use-ExcelWorkbook` は `Workbook.Close failed` と `Excel.Quit failed` の警告を出した。Excel のプロセスは残らなかった。

## 中断モードを VBE のリセットで戻す

コンパイルエラーのダイアログを閉じた後の中断モードは、VBE の「リセット」を別の runspace から実行すると終わり、`Run` が戻った（2026-11-16、同じ環境で 16 回）。`Invoke-ExcelMacro` は、この方法を使う。

手順は、上の実験の 4 を次に替えたもの。

1. ダイアログに `WM_CLOSE` を送り、閉じるのを待つ。1 秒待つ。
2. 見張りの runspace から、`VBE.CommandBars.FindControl(1, 228)` の `Execute()` を呼ぶ。
3. `VBE.MainWindow.Visible` を偽にする。
4. 元の runspace で、値を返すだけの別のマクロを `Run` で呼ぶ。

結果は次のとおり。

- ID 228 のコントロールの `Caption` は「リセット(&R)」で、中断モードでは `Enabled` が真だった。ID 186 は「継続(&C)」で真、ID 189 は「中断(&K)」で偽だった。
- `Execute()` の 6〜8 ms 後に、`Run` が例外 `0x800A9C68` になった。実行時エラーを「終了」で閉じた時と同じ番号である。
- リセットの 1 秒後、VBE のウィンドウの題名から「[中断]」が消えていた。`MainWindow.Visible` を偽にすると、ウィンドウは隠れた。
- 別のマクロは値を返した。ブックは閉じられ、Excel は `Quit` で終わり、`Use-ExcelWorkbook` の警告は無かった。リセットの直後に閉じた 6 回と、2 秒待ってから閉じた回のどちらも同じだった。
- 見張りの runspace が COM の参照を得る方法は 2 つ試し、どちらも通った。ブックのパスからの `Marshal.BindToMoniker`（13 回）と、元の runspace の `Excel.Application` をそのまま渡す形（3 回）。見張りの runspace のスレッドは MTA だった。`Run` で止まっている間も、もう一方からの COM の呼び出しは通る。
- リセットをしない対照の回は、上限の 40 秒で見張りのプロセスが Excel を止めるまで、`Run` が戻らなかった。

この実験で確かめたのは、宣言の無い変数のコンパイルエラーだけ。実験では、VBE のウィンドウが隠れるまで、ダイアログとあわせて 2〜3 秒、画面に出ていた。

## Invoke-ExcelMacro の扱い

見つけたダイアログを、次の順に見分ける。

1. ID 4800 のボタンがある: 実行時エラー。「終了」を押し、文言を例外にする。Excel は止めない。
2. VBE のウィンドウが表示されている: コンパイルエラー。`WM_CLOSE` で閉じ、「リセット」の `Enabled` が真になるのを待って実行し、VBE を隠す。文言を例外にする。Excel は止めない。固定の 1 秒は待たない。リセットで例外が出た時と、リセットの後 3 秒たっても `Run` が戻らない時は、その Excel を止める。`AccessVBOM` が無効な時の動きは確かめていないが、`VBE` を読めずに、この止める側に入る見込み。
3. ID 2 とヘルプ（ID 9）のほかのボタンがある: 答えを選ぶダイアログ（Yes/No、OK/Cancel の MsgBox など）。`WM_CLOSE` を送らず、文言を読んでその Excel を止める。
4. それ以外: `WM_CLOSE` を送る。3 秒で閉じれば先へ進み、文言を警告に出す。閉じなければ、文言を読んでその Excel を止める。

3 の判定は `XlWright.Dialog.AsksForAnswer` にあり、`Invoke-RibbonButton.ps1` も、リボンのボタンが出したダイアログに同じ判定を使う。止める時は、どちらも `XlWright.ExcelProcess.StopIfSame` で PID と起動時刻を照合する。

答えを選ぶダイアログでどれも選ばないのは、選ぶとマクロの動きをスクリプトが決めることになるため。`WM_CLOSE` は、MsgBox を ID 2（`IDCANCEL`）のボタンで閉じる。OK だけの MsgBox の OK は ID 2 なので、閉じても答えは 1 つしかない。OK/Cancel の MsgBox は、`WM_CLOSE` で Cancel を選んで閉じ、マクロは Cancel の側へ進んだ（下の、ボタンを見ない写し）。

### ボタンの ID

3 の判定が、閉じてよいダイアログを止める側に入れないかを確かめた（2026-11-17、同じ環境で各 1 回）。マクロが出したダイアログのボタンの ID を読み、`WM_CLOSE` を送った。

| ダイアログ | ボタンの ID | `WM_CLOSE` | 戻り値 |
| --- | --- | --- | --- |
| OK だけの MsgBox | 2 | 閉じた | 1（`vbOK`） |
| OK とヘルプの MsgBox（`vbMsgBoxHelpButton`） | 2、9 | 閉じた | 1 |
| Win32 の `MessageBoxW`（`MB_OK`、`Declare` で呼ぶ） | 2 | 閉じた | 1 |
| VBA の `InputBox` | 1、2、4902 | 閉じた | 空の文字列（キャンセル） |
| `Application.InputBox` | `#32770` が 8 秒で見つからない | - | 実験の見張りが Excel を止めた |

- OK だけのダイアログは、どれも OK が ID 2 だった。3 の判定で止める側に入るものは無かった。
- 表示した Excel で、マクロから `DisplayAlerts` を真にして Excel 自身の警告を出した（2026-11-17、各 1 回）。セルの結合（値のある 2 セル）とシートの削除の警告は、8 秒待っても `#32770` として見つからず、実験の見張りが Excel を止めた。警告のウィンドウのクラスは読んでいない。循環参照になる数式を入れたマクロは、警告を出さずに値を返した。Excel 自身の警告は 3 の判定まで届かないので、判定の変更で新しく止めるものは無い。見つけられない警告では、`Run` が上限まで戻らない。
- VBA の `InputBox` は、`WM_CLOSE` でキャンセルを選んで閉じた。3 の判定では止める側に入る。答えを選ぶダイアログとして扱われる。`Invoke-ExcelMacro` を通しては確かめていない。

`tests/Test-XlWright.ps1` の `macro` と `macro-dialog` が、これを確かめる（2026-11-16）。実験に加えて、次を確かめた。

- `Sub` の中の実行時エラーも、`Function` と同じに扱えた。その後に別のマクロが値を返した。
- OK だけの MsgBox を 2 回続けて出すマクロは、2 回とも閉じて値が返った。
- 引数を 3 つ渡して値が返った。`Sub` を呼ぶと、戻り値は無い。
- 宣言の無い変数と構文エラーのどちらのコンパイルエラーも、文言を入れた例外になった。同じマクロをもう一度呼んでも同じ例外になり、その後に別のモジュールのマクロが値を返した。VBE は隠れ、Excel は警告なしで終わった。
- Yes/No の MsgBox は 30 秒以内に例外になり、Excel は残らなかった。
- OK/Cancel の MsgBox も 30 秒以内に例外になり、Excel は残らなかった（2026-11-17）。
- `macro-dialog` は前面のウィンドウを 50 ms ごとに見る。Excel は一度も前面に来なかった。
- 「終了」の ID を変え、引数も渡さないようにした写しは、`macro` と `macro-dialog` の両方で落ちた。
- リセットを実行しない写しは、3 秒後に Excel を止め、`macro-dialog` と `vba-test-fail` で落ちた。
- VBE を隠さない写しも、同じ 2 項目で落ちた。この写しでは、後の MsgBox が 2 の扱いになった。VBE が表示されたままだと、MsgBox とコンパイルエラーを見分けられない。
- リボンのボタンが出した OK/Cancel の MsgBox も、閉じずに Excel を止めた（`ribbon`、2026-11-17）。判定を外した `Invoke-RibbonButton.ps1` の写しでは、× で閉じて `Clicked` が `true` になり、`ribbon` で落ちた。
- 起動時刻を照合しない `StopIfSame` の写しと、居ないプロセスを別のプロセスと返す写しは、`stop-own-process` で落ちた。`StopIfSame` は、止めた（`Stopped`）、居ない（`Gone`）、別のプロセスか起動時刻を読めない（`NotSame`）、止められなかった（`Failed`）を返し、呼び出し元はそれぞれ別の文言にする。
- 古い版の `XlWright.Dialog` を先に読み込んだ session では、module の読み込みが、新しい `powershell.exe` で流すよう求める例外になった（`stop-own-process`）。型は同じ session で置き換えられないので、`Revision` の値で見分ける。照合を外した写しは、読み込みが通り、`stop-own-process` で落ちた。
- ボタンの ID を見ない写しは、`macro-dialog` で落ちた。OK/Cancel の MsgBox が `WM_CLOSE` で閉じ、例外にならなかった（2026-11-17）。

## 未確認

- 宣言の無い変数と構文エラーのほかのコンパイルエラー。
- 呼ぶ前から VBE のウィンドウが表示されている時。`Invoke-ExcelMacro` は、その時は VBE を隠さない。
- `AccessVBOM` が無効な時のコンパイルエラー。
- 「デバッグ」や「継続」を押した時。
- 文書モジュールのマクロでの実行時エラー。
- `On Error` の無い実行時エラーが、イベント（`Workbook_Open` など）の中で出た時。`Invoke-ExcelMacro` が見張るのは、自分が呼んだマクロの実行中だけ。
- `DisplayAlerts` を真にしたマクロが出す Excel 自身の警告のウィンドウの種類と、それを見つけて閉じる方法。
- `Application.InputBox`、ファイルを選ぶダイアログ、UserForm。`Application.InputBox` は、8 秒待っても `#32770` が見つからなかった（下の「ボタンの ID」）。見張りは見つけられず、`Run` は上限まで戻らない見込み。
- ダイアログを見つけた時点で、VBE のウィンドウがまだ表示されていないコンパイルエラーがあるか。あれば 3 の扱いになり、`Run` は上限まで戻らない。実験と試験では、毎回表示されていた。
- 32-bit の Excel、日本語以外の表示言語（ボタンの ID が同じか）。
