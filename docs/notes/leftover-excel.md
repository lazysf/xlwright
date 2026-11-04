# 残った専用のExcelへの接続

設計の[残った専用のExcelの片付け](../architecture.md#残った専用のexcelの片付け)の未確認のうち、成立しない時に止めて人と相談するものを確かめた記録です。COMで起動したExcelが、起動したプロセスの強制終了の後も残るかと、残ったExcelにウィンドウから接続できるかの2つです。
確かめた日は2026-11-04です。

## 版

- Windows 11 Pro 10.0.26300、.NET 10.0.14（x64）
- Microsoft 365 Excel 64-bit 16.0.20527.20112
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`dotnet`

## 再現の手順

```powershell
dotnet run --project spikes\leftover -- leave .xlwright\spikes\leftover\a.json .xlwright\spikes\leftover\Leftover.xlsm
dotnet run --project spikes\leftover -- probe .xlwright\spikes\leftover\a.json
dotnet run --project spikes\leftover -- watch .xlwright\spikes\leftover\a.json 600
dotnet run --project spikes\leftover -- kill .xlwright\spikes\leftover\a.json
```

- `leave`は、STAのスレッドから専用のExcelを1つ起動し、所有を確かめます（新しいPID、`EXCEL`、起動を求めた後の生成）。ブックを渡した時は開きます。PIDと生成時刻をファイルに書き、COMを解放せずに自分のプロセスを強制終了します。サーバーが強制終了された時の代わりです。
- `probe`は、記録のPIDと生成時刻が一致するプロセスの`XLMAIN`と、その子の`EXCEL7`を列挙します。各ウィンドウに`AccessibleObjectFromWindow`（`OBJID_NATIVEOM`）を呼びます。接続できたら、`Application.Hwnd`のPIDを記録と照合し、ブックを保存せずに閉じてQuitし、75秒待ちます。
- `watch`は、記録のプロセスが終わるまで5秒ごとに照会します。`kill`は、記録と照合してから終了させます。
- ブックは、[段階1の流れ](stage1-flow.md)の主ブックの写しです。ブックを渡さない`leave`（`b.json`）も行いました。

## 判定

### 起動したプロセスの強制終了の後に残るか

支持（ブックのある時と無い時に1回ずつ）。どちらのExcelも、起動したプロセスの強制終了の後に残った。ブックの無いExcelは、10分経っても自ら終わらなかった。

### ブックのある残ったExcelへの接続

支持（1回）。非表示の`XLMAIN`が2つと、`EXCEL7`が1つあった。

- `XLMAIN`への呼出しは、どちらも`0x80004005`を返した。
- `EXCEL7`への呼出しは5 msで成功した。`Application.Hwnd`のPIDは記録と一致し、`Workbooks`に写しのブックが、`ProtectedViewWindows`に0が返った。
- ブックを閉じてQuitした後、プロセスは75秒では終わらなかった。Quitから約12分の後も残っていた。

### ブックの無い残ったExcelへの接続

不一致（2回）。ブックの無いExcelには`EXCEL7`が無く、`XLMAIN`への呼出しは`0x80004005`を返した。ウィンドウからは接続できなかった。

- 1回は、ブックを渡さない`leave`のExcel。
- 1回は、上の試行でブックを閉じてQuitした後も残ったExcel。

### 非表示のブックのウィンドウとProtected View

支持（1回ずつ）。どちらの形でも、`XLMAIN`の子に`EXCEL7`があった。そのため、`EXCEL7`が無いことで、ブックのウィンドウが無いと見分けられる。

```powershell
dotnet run --project spikes\leftover -- leave .xlwright\spikes\leftover\hidden.json .xlwright\spikes\leftover\Leftover.xlsm hidden
dotnet run --project spikes\leftover -- leave .xlwright\spikes\leftover\protected.json .xlwright\spikes\leftover\Leftover.xlsm protected
```

- `hidden`は、ブックを開いてから`Windows(1).Visible`を偽にする。`EXCEL7`が1つあり、そこから接続できた。`Workbooks`にブックが返った。閉じてQuitした後も、プロセスは75秒で終わらなかった。
- `protected`は、`ProtectedViewWindows.Open`で開く。`Workbooks`は0、`ProtectedViewWindows`は1だった。`EXCEL7`が2つあった。1つ目から得たオブジェクトには`Application`が無く、2つ目から得た`Application.Hwnd`は0だった。実験は接続を止め、Excelに触れなかった。

### 応答しない残ったExcelへの照会

支持（1回）。期限を付けた使い捨てのSTAで、照会を10秒で見切れた。

```powershell
dotnet run --project spikes\leftover -- leave .xlwright\spikes\leftover\busy.json .xlwright\spikes\leftover\Leftover.xlsm busy
dotnet run --project spikes\leftover -- probe .xlwright\spikes\leftover\busy.json 10
```

- `busy`は、ブックを開いた後に`Application.Run`で`Calc.Forever`（中身の無い`Do`〜`Loop`）を呼び、3秒後に別のスレッドから自分のプロセスを強制終了する。
- `probe`は、ウィンドウごとに使い捨てのSTAで`AccessibleObjectFromWindow`を呼び、10秒で見切る。`XLMAIN`の2つと`EXCEL7`の1つは、どれも10秒では戻らなかった。
- 見切った呼出しのうち2つは、後から`0x80004005`を返した。`XLMAIN`は30,023 ms、`EXCEL7`は10,008 msだった。

### 強制終了の後に残るもの

支持（2回）。ブックを開いた専用のExcelを強制終了した後、`~$Leftover.xlsm`は残らなかった。

- 1回は、上の`busy`のExcel。実行中は`~$Leftover.xlsm`があり、`kill`の後に無くなった。
- 1回は、開き直したExcel（`reopen.json`）。閉じてQuitした後も残ったため、`kill`で終了させた。
- 開き直すと、`Workbooks.Open`は約2秒で戻り、`ReadOnly`は偽だった。

## 後始末とGit差分

- 残った6つのExcel（PID 11096、18344、12392、8632、7124、3484）は、記録と照合してから`kill`で終了させた。
- `EXCEL.EXE`は、実験の前と後で0だった。
- 実験の後の`git status`は、TODO.mdの周期の追記と、`spikes/leftover/`の追加だけだった。

## 未確認

- 変更して自動回復の記録が作られた後に強制終了した主ブックを、次に開く時の影響。今回の試行は、どれもブックを変更していない。
- Quitの後もプロセスが残る理由。強制終了したプロセスが持っていた参照が残っていると見られるが、確かめていない。
