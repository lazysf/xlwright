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

## 後始末とGit差分

- 残った4つのExcel（PID 11096、18344、12392、8632）は、記録と照合してから`kill`で終了させた。
- `EXCEL.EXE`は、実験の前と後で0だった。
- 実験の後の`git status`は、TODO.mdの周期の追記と、`spikes/leftover/`の追加だけだった。

## 未確認

- 応答しない残ったExcelへの照会を、時間で打ち切れるか（設計の未確認の4つ目）。実施していない。
- Quitの後もプロセスが残る理由。強制終了したプロセスが持っていた参照が残っていると見られるが、確かめていない。
