# 人が開くブックと表示要求

設計の未確認のうち、人の操作が要る2つを確かめた記録です。人がエクスプローラーから開いたブックが残った専用のExcelに入るか（[日常のExcelとの分離](../architecture.md#日常のexcelとの分離)）と、非表示の残ったExcelをウィンドウの表示要求で人へ見せられるか（[人への引継ぎ](../architecture.md#人への引継ぎ)）です。
確かめた日は2026-11-04です。

## 版

- Windows 11 Pro 10.0.26300、.NET 10.0.14（x64）
- Microsoft 365 Excel 64-bit 16.0.20527.20112。`VBAWarnings`は2
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`dotnet`

## 再現の手順

```powershell
dotnet run --project spikes\leftover -- leave .xlwright\spikes\handoff\a.json .xlwright\spikes\handoff\Handoff.xlsm
dotnet run --project spikes\leftover -- list .xlwright\spikes\handoff\a.json
# 人が、エクスプローラーから .xlwright\spikes\handoff\Human.xlsm をダブルクリックし、画面を見て閉じる
dotnet run --project spikes\leftover -- list .xlwright\spikes\handoff\a.json
dotnet run --project spikes\leftover -- show .xlwright\spikes\handoff\a.json
# 人が、表示されたExcelを見て、保存せずに閉じる
dotnet run --project spikes\leftover -- watch .xlwright\spikes\handoff\a.json 90
dotnet run --project spikes\leftover -- list .xlwright\spikes\handoff\a.json
dotnet run --project spikes\leftover -- kill .xlwright\spikes\handoff\a.json
```

- `leave`は、[残った専用のExcelへの接続](leftover-excel.md)と同じく、ブックを開いた専用のExcelを残して自分のプロセスを強制終了します。
- `list`は、記録と照合したExcelに`EXCEL7`から接続し、`Workbooks`と`Visible`を読んで、閉じずに参照を解放します。
- `show`は、COMを使わずに、`EXCEL7`を持つ`XLMAIN`に`ShowWindow`（`SW_SHOW`）と`SetForegroundWindow`を呼びます。
- 2つのブックは、同じ内容の写しです。

## 判定

### 人がエクスプローラーから開いたブック

支持（1回）。非表示の残った専用のExcel（PID 18448）がある時に、人がエクスプローラーから開いたブックは、新しいExcel（PID 12360）の表示されたウィンドウで開いた。

- 新しいExcelのコマンドラインは、ブックのパスだけで、`/automation`を含まなかった。
- その間も、専用のExcelの`Workbooks`は`Handoff.xlsm`だけだった。
- マクロは、セキュリティの警告のバーで無効になった。人は有効にせずに閉じ、12360は終わった。
- 日常のExcelが動いていない時の1回だけです。

### 表示要求で見せられるか

支持（1回）。`show`の後、`XLMAIN`は表示状態になり、`SetForegroundWindow`は真を返した。人は、`Handoff.xlsm`のExcelが見えることを確かめ、保存せずに閉じた。

- 表示の前の`Application.Visible`は偽だった。`show`はこの値を変えていない。

### 人が閉じた後にプロセスが残るか

残った（1回）。人が閉じた後、90秒の照会の間、18448は終わらなかった。

- その後の照会で、`XLMAIN`は2つとも非表示で、`EXCEL7`は無かった。ブックを閉じた後に残ったExcelと同じ形で、ウィンドウからは接続できなかった。
- 人がいつ閉じたかは記録していない。人の返事の時点（11:09:51）から照会を始めた。

## 後始末とGit差分

- 18448は、記録と照合してから`kill`で終了させた。
- `EXCEL.EXE`は、実験の前と後で0だった。2つのブックのハッシュは、実験の前と後で変わらなかった。
- 実験の後の`git status`は、TODO.mdの周期の追記と、`spikes/leftover/Program.cs`の変更だけだった。

## 未確認

- 日常のExcelが動いている時に、人がエクスプローラーから開くブックの行き先。
- 人が表示されたExcelで主ブックを開く時に出るダイアログ。
- 人が閉じた後に残るプロセスが、自ら終わるまでの時間。
