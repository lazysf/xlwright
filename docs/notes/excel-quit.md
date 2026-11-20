# Quit の後に Excel が終わるまで

- Date: 2026-11-20
- 出典: 一時フォルダの実験用スクリプトの実測と、同じ時間帯の Windows のイベントログ（Application、OAlerts、System）、Excel の診断ログ（`%TEMP%\Diagnostics\EXCEL\`）、`%LOCALAPPDATA%\Microsoft\Office\Licenses\` のファイルの時刻。Microsoft 365 Excel 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818（64-bit）。人は何も触らない状態で測った。

`Stop-ExcelSession`（`plugins/xlwright/skills/xlwright/scripts/XlWright.psm1`）は、ブックを閉じて `Quit` した後、Excel が終わるのを 5 秒まで待つ。終わらなければ、この呼び出しが起動した Excel を PID と起動時刻で照合して止める。Excel はふつう `Quit` の約 2.5 秒後に自分で終わり、5 秒で足りる。終わらない Excel は、測った範囲では 60 秒待っても終わらなかったので、待ちを延ばしても止める回が減るだけで、速くはならない。

## 実験の手順

1. `plugins/xlwright/skills/xlwright/scripts/` を一時フォルダに写した。写しの `Stop-ExcelSession` は、`Quit` の後に 100 ms ごとにプロセスを見て、終わるまでの時間を記録する。60 秒たっても終わらなければ、今と同じく照合して止める。
2. 1 回ごとに新しい `powershell.exe` で、次の 6 種類の操作のどれかを 1 つ行った。`Use-ExcelWorkbook` で開くだけ、`Read-ExcelRange` で読む、`Invoke-ExcelMacro`、`Test-ExcelVbaCompile`、`$xl.Run`、`New-ExcelWorkbook`。ブックは `Util.bas` を取り込んだ `.xlsm` を使った。
3. 6 種類を 1 巡として 6 巡流した。続けて、6 分 Excel を使わずに置いてから開くだけを 8 回流し、その後にまた 6 種類を 6 巡流した。

## 結果

| 流し方 | 回数 | 2.2〜2.9 秒で終わった | 5〜7 秒で終わった | 60 秒たっても終わらない |
| --- | --- | --- | --- | --- |
| 6 種類を 6 巡（11:27〜11:31） | 36 | 32 | 2 | 2 |
| 6 分置いてから開くだけを 8 回（11:38） | 8 | 8 | 0 | 0 |
| 6 種類を 6 巡（11:39〜11:42） | 36 | 36 | 0 | 0 |

- 遅れた 4 回は、どれも最初の 1 巡に出た。開くだけ（5.6 秒）、読む（終わらない）、`Invoke-ExcelMacro`（6.9 秒）、`Test-ExcelVbaCompile`（終わらない）の 4 回。同じ操作は、ほかの巡では 2.2〜2.9 秒で終わったので、遅れるかは操作の種類では決まらない。
- 6 分 Excel を使わずに置いても、遅れる回は出なかった。
- [dev-cycle-timing.md](./dev-cycle-timing.md) と [vba-one-session.md](./vba-one-session.md) の実測でも、`Quit` の後 5 秒で終わらない回があった（10:51 に 2 回、10:56 と 10:57 に 1 回ずつ）。この 4 回の Excel の診断ログは、古いものから消されていて残っていなかった。

## 終わらなかった Excel が起動の時にしていたこと

Excel は起動ごとに、`%TEMP%\Diagnostics\EXCEL\` に診断ログを 1 つ書く。ファイル名の数字は起動時刻（1970 年からのナノ秒）で、記録は起動から約 1 秒で終わる。終了の場面は記録されないので、起動の時に出たイベントを比べた。

| イベント | 終わらなかった 2 回 | 5〜7 秒で終わった 2 回 | 同じ時間帯のほかの起動 77 回 |
| --- | --- | --- | --- |
| `Office.Licensing.Nul.Validation.FullValidation`（ライセンスの完全な検証。3268 ms と 4192 ms、成功） | 2 | 0 | 0 |
| `Office.Licensing.TelemetryFlow.FetchLicenseFlow`（`Data.Source` が `RenewLicense`、`Data.IsRenewal` が `true`） | 2 | 0 | 0 |
| `Office.Docs.DocumentMru.AggregatedMru.SendRecentServiceReadRequest` | 2 | 0 | 0 |
| `Office.Privacy.OffersOptIn` | 2 | 0 | 0 |

- 77 回は、11:26〜11:43 に起動した Excel の診断ログのうち、遅れた 4 回のほかのすべて。76 回は 2.2〜2.9 秒で終わった計測の回で、残りの 1 回は計測と対応づけていない。
- 同じイベントは、11:17〜11:23 に試験を流した時の起動の多くにも出ていた（`FullValidation` は 46 回）。同じ時間帯の起動には、前の Excel が正しく終わらなかったことを示す `Office.System.SystemHealthUngracefulAppExitDesktop` が 17 回出ていた。この時間帯の Excel のどれが `Quit` の後に終わらなかったかは、測っていない。
- 11:29 より後の起動には、これらのイベントは出なかった。
- `%LOCALAPPDATA%\Microsoft\Office\Licenses\` のライセンスのファイルは、11:22:46 に書き換わっていた。その後の 11:27 と 11:28 の起動でも、更新（`RenewLicense`）は始まっていた。
- 10:53:22〜10:53:31 に、Windows Installer が入っている MSI の製品を 157 件再構成していた（`MsiInstaller` の ID 1035、どれも成功）。`Win32_Product` の問い合わせが起こしたもので、Office の Click-to-Run の部品 2 つも含まれていた。10:51 の 2 回はこれより前に遅れているので、遅れの原因はこれだけではない。その後の遅れに影響したかは分からない。
- Windows のイベントログには、遅れた時刻に重なる Office の警告やエラーは無かった。OAlerts には Excel の起動ごとに ID 300 が 1 件ずつあるだけで、遅れた回とほかの回で違いは無かった。ライセンスの保護のサービス（`Microsoft-Windows-Security-SPP`）のイベントは 10:51 と 11:41 にあったが、11:41 の前後の起動は遅れなかった。

ここから、Office が起動の時にライセンスの更新を始めた Excel は、`Quit` の後も終わらないことがある、と見ている。更新がいつ、なぜ始まるかは、ログからは分からない。ライセンスの状態はこちらから変えない（`AGENTS.md`）ので、更新を起こして確かめることはしていない。

## 今の形にしている理由

- 待ちを 5 秒より延ばしても、終わらない Excel は止めることになる。10 秒にした場合、今回の 4 回では、5〜7 秒で終わった 2 回が自分で終わるのを待つ分だけ 0.5〜1.8 秒遅くなり、終わらない 2 回は 5 秒遅くなる。
- 止めるのは、ブックを閉じて `Quit` した後に限る。止めた後の次の起動は、前の Excel が正しく終わらなかったことを診断ログに記録する（`UngracefulAppExitDesktop`）。起動の失敗や回復の表示は見ていない。

## 確かめていないこと

- ライセンスの更新が始まる条件と、更新がどこまで進むと Excel が終わるか。
- ライセンスの更新の途中で Excel を止めることが、更新の完了やライセンスの状態に影響するか。更新が続けて起きた時間帯（11:17〜11:23）と、止めた回が重なっていたかは調べていない。
- 5〜7 秒で終わった 2 回が遅れた理由。この 2 回の診断ログには、ほかの回に出ないイベントは無かった。
- 終わらない間の Excel の状態（ウィンドウ、スレッド、通信）。写しのスクリプトはこれを 5 秒ごとに記録するが、記録を入れた後の 44 回には終わらない回が出なかった。
