# .NETのSTAからの専用のExcelの起動・所有・終了とCOMの呼出し

[設計](../architecture.md)の段階1の未確認のうち、主ブックを使わずに確かめられるものの記録です。
製品のサーバーと同じ.NETの専用のSTAから、`Excel.Application`を遅延バインド（`dynamic`）で扱いました。

## 版

- Windows 11 Pro 10.0.26300、.NET 10.0.14（x64）
- Microsoft 365 Excel 64-bit 16.0.20527.20112
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`dotnet`

## 再現の手順

```powershell
dotnet build spikes\sta-com -c Release
dotnet spikes\sta-com\bin\Release\net10.0-windows\StaCom.dll all
```

[実験](../../spikes/sta-com/Program.cs)は、試行ごとに専用のSTAで次を行います。追加の依存はありません。

1. 実行の前からあるEXCELのプロセスを記録する。
2. メッセージフィルターを登録し、時刻を記録してから`Activator.CreateInstance`で起動する。
3. `Hwnd`から`GetWindowThreadProcessId`でPIDを引き、そのPIDのハンドルを開く。同じハンドルで、名前が`EXCEL.EXE`であること、記録した時刻より後に生成されたこと、実行の前に無かったPIDであることを確かめ、コマンドラインを読む。確かめられなければ、そのExcelに触らずに止まる。
4. 試行ごとの操作の後、ブックと保護ビューのウィンドウが0であることを確かめて`Quit`を求める。強制終了は、所有を確かめた時に開いたハンドルでだけ行う。
5. 実行の前からあったプロセスが変わらず、実行が起動したEXCELが残っていないことを照合する。

試行は次の4つです。

- `start-stop`: 起動、所有の確定、`Quit`を3回。
- `retry-later`: スレッドAが`Application.Wait`（8秒）か、重い数式（100セル）の入力で待つ間に、別のSTAのスレッドBが`Application.Version`を呼ぶ。Bにメッセージフィルターがある時と無い時を比べる。フィルターは`RETRYLATER`を10秒まで再試行し、`REJECTED`は再試行しない。数式は、ファイルの無い新しいブックに入れ、保存せずに閉じた。
- `release`: `Quit`を呼ばずにCOMの参照をすべて解放し、30秒待つ。`UserControl`が偽のままの時と、真にした時を比べる。真にした時に`Quit`で終わるかも確かめる。
- `pending-kill`: スレッドAが`Application.Wait`（60秒）で待つ間に、別のスレッドから`TerminateProcess`で終了させる。

実行の前に、EXCELのプロセスは無かった。

## 判定

### COMから起動したEXCELのコマンドライン（[所有の確定](../architecture.md#所有の確定)）

支持。コマンドラインを表示した13回の起動のすべてで、コマンドラインは`"…\EXCEL.EXE" /automation -Embedding`だった。
コマンドラインは、`PROCESS_QUERY_LIMITED_INFORMATION`で開いたハンドルから、`NtQueryInformationProcess`の`ProcessCommandLineInformation`（60）で読めた。

### Quitの後、プロセスが終わるまでの時間（[終了](../architecture.md#終了)）

ブックを開いていない時は、支持（測った）。`Quit`は0〜2 msで戻り、プロセスは戻ってから2,218〜2,704 msで終わった（11回、終了コードはどれも0）。主ブックを開いて保存した後の時間は、測っていない。

不一致を1つ見つけた。新しいブックで`workbook.Worksheets[1].Range[…]`のように中間のオブジェクトを変数に取らずに使うと、`Quit`の後75秒経ってもプロセスが終わらなかった（2回）。その後に所有を確かめたハンドルで終了させた。`Quit`の後に`GC.Collect`と`GC.WaitForPendingFinalizers`を呼んでも、変わらなかった。中間のオブジェクトを変数に取って`Marshal.FinalReleaseComObject`で解放すると、2,388 msで終わった。

### `SERVERCALL_REJECTED`、`RETRYLATER`、メッセージフィルター（[構成](../architecture.md#構成)、[期限](../architecture.md#期限と大きさの上限)）

- 起動の直後の`SERVERCALL_REJECTED`: 不一致（起きなかった）。メッセージフィルターを登録したスレッドからの13回の起動で、起動の直後の呼出しに`RetryRejectedCall`は一度も呼ばれなかった。この端末では、起動の時にサインインのダイアログは出ていない。
- `Application.Wait`の間の呼出し: Excelは受け付けた（Bは1〜3 msで戻った）。
- 再計算の間の呼出し: `RETRYLATER`が返った。フィルターのあるBは、`RETRYLATER`を3〜4回受けた後、1,376〜1,654 msで成功した。`RetryRejectedCall`が呼ばれる間隔は、100 msの再試行を返しても約500 msだった。フィルターの無いBは、約1秒後に`0x80010001`（`RPC_E_CALL_REJECTED`）で失敗した。遅延バインドでは、`GetIDsOfNames`の失敗として`MissingMemberException`に包まれて返った。

メッセージフィルターによる`RETRYLATER`の再試行は、実Excelで働いた。
保存と`Close`の遅延した戻りは、ブックを保存していないため確かめていない。

### COMの参照の解放で、専用のExcelが自ら終わるか（[起動時の設定](../architecture.md#起動時の設定)）

- `UserControl`が偽（起動の直後の値）のまま、参照をすべて解放した時: 終わった（3回）。解放から約2.5秒だった。
- 所有を確定した直後に`UserControl`を真にしてから解放した時: 30秒経っても終わらなかった（3回）。その後に所有を確かめたハンドルで終了させた。
- `UserControl`を真にしてから`Quit`を呼んだ時: 2,704 msで終わった（1回）。

実験のプロセスが失敗で終わった時も、参照が解放されて、ブックの無いExcelは自ら終わった（1回）。

### 保留中のCOM呼出しがある時の終了（[状態と結果不明](../architecture.md#状態と結果不明)）

支持。`Application.Wait`の60秒の呼出しの途中で終了させると、呼出しは18〜20 ms後に`0x800706BE`（`RPC_S_CALL_FAILED`）で戻った（2回）。スレッドAには、メッセージフィルターを登録していた。
マクロの実行中に終了させた時の戻り方は、確かめていない。

## 後始末とGit差分

- 正常に終わった実行の後は、どれもEXCELのプロセスが0だった。実行の前からあったプロセスは無かった。
- 実験の不具合（スレッドBの捕まえていない例外）で、実験のプロセスが新しいブックを開いたExcel（PID 21940）を残して終わったことが1回あった。記録したPIDと生成時刻（100ナノ秒の単位まで）が一致することを確かめてから、終了させた。終了させた後に、EXCELのプロセスが0であることを確かめた。この時だけは、所有を確かめた時のハンドルではなく、PIDで開き直したハンドルで終了させた。
- ブックのファイルは開いていない。新しいブックは保存せずに閉じた。
- Gitの差分は、実験と本書の追加と、設計とTODO.mdの更新だけです。
