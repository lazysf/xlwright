# ブックの保存・閉じる・開くと、`AccessVBOM`が0の時

[設計](../architecture.md)の段階1の未確認のうち、VBAのプロジェクトを扱えないまま確かめられるものの記録です。
[.NETのSTAからの起動と終了](sta-com.md)と同じ実験に、試行`workbook`を加えました。

## 版

- Windows 11 Pro 10.0.26300、.NET 10.0.14（x64）
- Microsoft 365 Excel 64-bit 16.0.20527.20112
- Excelの設定（`HKCU\Software\Microsoft\Office\16.0\Excel\Security`）: `AccessVBOM`は0、`VBAWarnings`は2。同じキーのポリシーの値は無かった。
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`dotnet`

## 再現の手順

```powershell
dotnet build spikes\sta-com -c Release
dotnet spikes\sta-com\bin\Release\net10.0-windows\StaCom.dll workbook
```

試行は、所有を確定した専用のExcelで、次を3回繰り返します。所有の確定と後始末の照合は、[前の記録](sta-com.md#再現の手順)と同じです。

1. 所有を確定した直後に`UserControl`を真にする。
2. 新しいブックの`A1`に値を、`B1`に`=A1*21`を入れ、`Calculate`の後に`CalculationState`と`B1`を読む。
3. `.xlwright/spikes/workbook/Spike.xlsm`（Git管理外）へ`SaveAs`し、閉じる。
4. `AutomationSecurity`を`Low`、`EnableEvents`を偽にして開き直し、`ReadOnly`と`FullName`を確かめる。
5. `Workbook.VBProject`を取得する。
6. 2と同じ書込みと計算をして`Save`し、閉じる。`EnableEvents`を戻し、`Quit`して終了を待つ。

使った中間のCOMのオブジェクトは、すべて変数に取り、`Marshal.FinalReleaseComObject`で解放しました。

## 判定

### `AccessVBOM`が0の時のエラーの値（[VBAの反映](../architecture.md#vbaの反映)）

支持（値を確かめた）。`Workbook.VBProject`の取得は、`0x800A03EC`で失敗した（3回）。メッセージは「プログラミングによる Visual Basic プロジェクトへのアクセスは信頼性に欠けます」だった。
`0x800A03EC`は、Excelがさまざまな失敗で返す値です。値だけでは、`AccessVBOM`が原因だと見分けられません。

### 保存とCloseの戻り（[構成](../architecture.md#構成)）

小さなブックでは、遅延した戻りは起きなかった。`RetryRejectedCall`は一度も呼ばれなかった。

- `SaveAs`: 56〜121 ms
- `Save`: 18〜22 ms
- `Close`: 7〜23 ms
- `Open`: 46〜70 ms

大きなブックでは確かめていない。

### 保存した後のQuitから終了までの時間（[終了](../architecture.md#終了)）

支持（測った）。保存して閉じた後、`Quit`は1 ms以内に戻り、プロセスは2,284〜2,602 msで終わった（3回、終了コードは0）。`UserControl`は真だった。

### `Calculate`の後の`CalculationState`と値（[起動時の設定](../architecture.md#起動時の設定)）

支持。`Calculate`の後の`CalculationState`は、6回とも`xlDone`（0）だった。`B1`は、入れた値に応じて42と63だった。自動計算のままで、計算の方法は変えていない。

### `AutomationSecurity=Low`と`EnableEvents=False`で開けるか（[起動時の設定](../architecture.md#起動時の設定)）

開けた。`ReadOnly`は偽で、`FullName`は保存したパスと一致した。
ブックにVBAが無いため、マクロの実行と、起動のイベントが動かないことは確かめていない。

## 後始末とGit差分

- 各回の後に、EXCELのプロセスは0だった。実行の前からあったプロセスは無かった。
- `Spike.xlsm`は、各回の後に消した。実行の後、`.xlwright/spikes/workbook/`は空だった。
- Gitで管理するブックには触れていない。
- Gitの差分は、実験の試行の追加と、本書の追加と、設計とTODO.mdの更新だけです。
