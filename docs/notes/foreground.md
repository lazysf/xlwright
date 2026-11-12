# リボンの操作で Excel を前面に出さない

- Date: 2026-11-12
- 出典: 前面のウィンドウ（`GetForegroundWindow`）を 50〜100 ms ごとに記録しながら、`tests/Test-XlWright.ps1` と一時フォルダの実験用スクリプトを流した実測。Excel 16（Microsoft 365）、Windows 11、Windows PowerShell 5.1、画面の拡大率 100%。特に書かない限り、Claude のウィンドウを前面にして、人は何も触らない状態で測った。

人がスキルの作業中に Excel 以外のアプリを使えるように、リボンのタブの切り替えとボタンの押下を MSAA の既定の動作（`IAccessible.accDoDefaultAction`）で行っている。実装は `plugins/xlwright/skills/xlwright/scripts/Ribbon.psm1` の `Select-RibbonTab` と `XlWright.RibbonMsaa` を正本とする。

## なぜ MSAA の既定の動作か

UI Automation の `SelectionItemPattern.Select()`（タブ）と `InvokePattern.Invoke()`（ボタン）は、Excel を前面に出す。前面にいる間のキー入力は Excel に入る。同じ要素の MSAA の既定の動作（タブは「切り替え」、ボタンは「クリック」）は、前面を変えない。

| 操作 | 方法 | Excel が前面にいた割合 |
| --- | --- | --- |
| タブの切り替え | UI Automation の `Select()` | 15/15 |
| タブの切り替え | MSAA の既定の動作 | 0/15 |
| ボタンの押下と MsgBox の表示中 | UI Automation の `Invoke()` | 押下後 14/15、表示中 5/5、閉じた後 10/10 |
| ボタンの押下と MsgBox の表示中 | MSAA の既定の動作 | すべて 0 |

MSAA で押しても、ボタンのマクロは動き、MsgBox の文言も UI Automation で読めた。MsgBox を閉じる `WindowPattern.Close()` は前面を変えない。

画面に出すこと自体（`Use-ExcelWorkbook -Visible`）と、配布ブックの出力（`Lock-VbaProject.ps1`）は前面を変えない。配布ブックの出力だけを流した時、前面は一度も移らなかった。

人が Windows Terminal で文字を打ち続けている間に `-Full -Case ribbon,ribbon-layout` を流した時も、前面は最後まで Windows Terminal のままで、試験は合格した。打った文字に欠けは無かった（人の目視）。

`tests/Test-XlWright.ps1` の `ribbon`、`ribbon-layout`、`publish` は、前面を 50 ms ごとに見て、Excel が一度も前面に来ないことを確かめる。タブの切り替えを `Select()` に戻した写しは `ribbon-layout` で、ボタンの押下を `Invoke()` に戻した写しは `ribbon` で、この確認に落ちた。

## 効かなかった方法

- 奪われた前面を、`AttachThreadInput` で前面のスレッドに入力をつないでから `SetForegroundWindow` で戻す: 毎回 `false` が返った。Windows は、前面に無いプロセスが前面を変えることを制限している。
- UI Automation の `AutomationElement.SetFocus()` で元のウィンドウにフォーカスを戻す: 毎回例外になった。
- 画面に出す間 `LockSetForegroundWindow` で前面の変更を禁じる: ロックは成功したが、タブを選ぶと Excel が前面に出た。
- Excel のウィンドウに `WS_EX_NOACTIVATE` を付ける: 前面に出た。
- 前面を奪われた後に Excel を隠して出し直す: 前面に出たままだった。
- WMI で起動した（前面を変える権利を持たない）プロセスからタブを選ぶ: 前面に出た。

## 採らなかった方法

- タブの位置へマウスのメッセージ（`WM_LBUTTONDOWN` / `UP`）を `PostMessage` で送る: 前面は変わらなかったが、座標に頼る。
- キー入力の偽装（`Alt` を送るなど）: 人の入力と混ざる。
- Windows の前面ロックの設定（`SPI_SETFOREGROUNDLOCKTIMEOUT`）を変える: システムの設定を変える。

## 未確認

- toggleButton など、普通のボタン以外の押下。
- 画面の拡大率が 100% 以外の時。
- 人が自分の Excel を使っている時。スキルの前提で、作業中は Excel を使わないことにしている。
