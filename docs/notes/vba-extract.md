# ForceDisableで開いたブックからの取り出し

段階4の`vba_extract`が依存する、設計の未確認の記録です（[VBAの反映](../architecture.md#vbaの反映)）。
確かめた日は2026-11-05です。

## 版

- Windows 11 Pro 10.0.26300、.NET 10.0.14（x64）
- Microsoft 365 Excel 64-bit 16.0.20527.20112
- Excelの設定: `AccessVBOM`は1、`VBAWarnings`は2
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`dotnet`

## 再現の手順

```powershell
dotnet run --project spikes\extract
```

実験は、STAのスレッドから専用のExcelを1つ起動し、所有を確かめます（新しいPID、`EXCEL.EXE`、起動を求めた後の生成）。確かめられなければ、触らずに止まります。
ブックは、実行ごとに作る一時フォルダにだけ作ります。専用のExcelのダイアログを200ミリ秒ごとに探し、文字を表示して`WM_CLOSE`で閉じます。1つの手順が30秒で戻らなければ、所有を確かめたプロセスを終了させます。

1. `AutomationSecurity=Low`・`EnableEvents=False`で、元のブック`Existing.xlsm`を作る。中身は、`Workbook_Open`・`Workbook_SheetCalculate`、シートの`Worksheet_Change`・`Worksheet_Calculate`、標準モジュールの`Auto_Open`と揮発性のUDF`Mark`（A1の`=Mark()`）、クラスモジュール、日本語のコメント、Scripting Runtimeの参照設定。イベントとUDFは、ブックの隣の印のファイルに1行を足す。
2. 対照として、元のブックを`Low`・`EnableEvents=True`で開き、`CalculateFull`、セルの書込み、`Application.Run "'Existing.xlsm'!Udf.Mark"`を行う。
3. 元のブックをアプリのフォルダに写し、写しを`AutomationSecurity=ForceDisable`・`EnableEvents=True`で開き、2と同じことを行う。イベントの抑止に頼らず、`ForceDisable`だけを確かめるためです。続けて、各モジュールを`CodeModule.Lines`で読み、`VBComponent.Export`で書き出し、参照設定を一覧する。
4. 写しを開いたまま`AutomationSecurity`を`Low`に戻し、2と同じことを行う。
5. Excel 4.0のマクロシートを加えた新しいブックを`.xlsm`で保存し、2つのブックのパッケージの`xl/`の下を一覧する。

## 判定

### ForceDisableで、イベントとUDFが動かないか

支持。`ForceDisable`で開いた写しでは、開く時も、`CalculateFull`・セルの書込み・`Application.Run`の後も、印のファイルは1つも作られなかった。
A1は、`CalculateFull`の後に`#NAME?`になった。`Application.Run`は`0x800A03EC`で失敗し、本文は、このブックでマクロを使えないか、すべてのマクロが無効になっているというものだった。ダイアログは出なかった。

対照の`Low`では、開いた時に`Workbook_Open`・UDF・シートとブックの計算のイベントが動き、`CalculateFull`・セルの書込み（`Worksheet_Change`）・`Application.Run`（戻り値7）でも動いた。`Auto_Open`は、`Low`でも動かなかった。外部のCOMの`Workbooks.Open`では動かないため（[記録](test-copy.md#起動のイベントの抑止)）。

### 開いた後にAutomationSecurityを戻した時

支持。写しを開いたまま`Low`に戻しても、印のファイルは作られず、A1は`#NAME?`のままだった。`Application.Run`は`0x800A03EC`で失敗し、本文は、セキュリティの設定でマクロが無効にされたため、開き直す必要があるというものだった。
マクロを止めるかは、開いた時の`AutomationSecurity`で決まった。

### ForceDisableで、VBAを取り出せるか

支持。`ForceDisable`の写しでも`VBProject`に触れられ、4つのモジュールのコードを`CodeModule.Lines`で読めた。種類は、`ThisWorkbook`と`Sheet1`が100、標準が1、クラスが2だった。標準とクラスのモジュールの日本語は、`Lines`の文字列で正しく読め、`Export`のファイルはCP932で読むと正しかった。
参照設定は、`VBA`・`Excel`・`stdole`・`Office`・`Scripting`の5つの`Name`・`Guid`・版・`BuiltIn`・`IsBroken`を読めた。

### Excel 4.0のマクロシートの見分け方

支持。マクロシートを加えたブックのパッケージには`xl/macrosheets/`があり、VBAだけのブックには無かった。マクロシートだけのブックには、`xl/vbaProject.bin`が無かった。
設計のとおり、取り出す前に`xl/macrosheets/`で拒否する。

## 分かったこと

- `ForceDisable`で開いたブックで再計算すると、UDFを使うセルは`#NAME?`になる。取り出しに使う写しは、保存せずに閉じる。

## 未確認

- UserFormのモジュールの取り出し。
- `ForceDisable`がExcel 4.0のマクロを止めるか。

## 後始末とGit差分

- 2回実行し、どちらも前後の`EXCEL.EXE`は0で、所有を確かめたプロセスだけを終えた。1回目は対照の2が無く、2回目に加えた。
- 一時フォルダは、どちらも消えた。
- Gitの差分は、`spikes/extract/`と文書だけ。管理するブックに触れていない。
