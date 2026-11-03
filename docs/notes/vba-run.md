# VBAの反映とマクロの実行

[設計](../architecture.md)の段階1の未確認のうち、VBAの反映とマクロの実行の記録です。
[.NETのSTAからの起動と終了](sta-com.md)と同じ実験に、試行`vba`を加えました。VBIDEは、PIAを使わずに遅延バインド（`dynamic`）で扱いました。

## 版

- Windows 11 Pro 10.0.26300、.NET 10.0.14（x64）
- Microsoft 365 Excel 64-bit 16.0.20527.20112
- Excelの設定: 人が「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にした後、`AccessVBOM`は1、`VBAWarnings`は2（警告を表示してすべてのマクロを無効にする）。
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`dotnet`

## 再現の手順

```powershell
dotnet build spikes\sta-com -c Release
dotnet spikes\sta-com\bin\Release\net10.0-windows\StaCom.dll vba
```

[試行](../../spikes/sta-com/VbaTrials.cs)は、`.xlwright/spikes/vba/App/`（Git管理外）にアプリのフォルダの形でソースを書き、主ブック`App.xlsm`を作ります。所有の確定と後始末の照合は、[前の記録](sta-com.md#再現の手順)と同じです。Excelの呼出しのうち止まりうるものは、時間を過ぎると、所有を確かめた時のハンドルで専用のExcelを終了させる見張りの下で行いました。

ソースは次の4つです。改行はCRLFです。

- `src/modules/Calc.bas`（CP932）: 日本語のコメント、`VB_Description`の属性を付けた関数、引数と戻り値の関数、日本語の文字列を返す関数、クラスを使う関数、セルに書くマクロ、0で割るマクロ、終わらないマクロ。
- `src/modules/CharsetUtf8.bas`（BOMの無いUTF-8）: 日本語の文字列を返す関数。
- `src/classes/Counter.cls`（CP932）: クラスのヘッダー、`VB_PredeclaredId = True`、既定のメンバー（`VB_UserMemId = 0`）。
- `Broken.bas`（CP932、`src/`の外）: `Option Explicit`で宣言の無い変数を使うマクロ。

試行は、専用のExcelを5回起動して行います。

1. 主ブックを作り、標準・クラスモジュールをすべて消してから`src/`のファイルをすべて`VBComponents.Import`する全置換を、2回行う。モジュールを書き出して属性を確かめる。試行の準備として、`ThisWorkbook`に、`D1`に書く`Workbook_Open`を足す。保存して閉じる。
2. `AutomationSecurity=Low`・`EnableEvents=False`で開き、`D1`を読む。`'App.xlsm'!`で修飾した`Application.Run`でマクロを呼ぶ。開いていないブックで修飾した名前も呼ぶ。閉じて、`EnableEvents=True`で開き直し、`D1`を読む。
3. 0で割るマクロを実行する（見張りは20秒）。
4. `Broken.bas`をImportし、そのマクロを実行する（見張りは20秒）。
5. 終わらないマクロを実行する（見張りは5秒）。

これを3回実行しました。3回目は、見張りが終了させる前に、そのプロセスの表示されたウィンドウのクラス名・有効状態・題名を記録しました。

## 判定

### Import（[VBAの反映](../architecture.md#vbaの反映)）

- 遅延バインドでのImport: 支持。`VBComponents.Import`は、標準・クラスモジュールを取り込めた。
- `VBAWarnings`が2の時のImport: 支持。設定を変えずに取り込めた。
- Importが使う文字コード: システムのANSIコードページ（CP932）。CP932のファイルの`"日本語"`は正しく返った。UTF-8のファイルの`"日本語"`は、別の6文字に化けた。
- 全置換と属性: 支持。2回の全置換のどちらの後も、モジュールは`Calc`・`CharsetUtf8`・`Counter`で、名前に番号が付かなかった。`Remove`はすぐに効いた。書き出したモジュールに、`VB_Description`・`VB_PredeclaredId = True`・`VB_UserMemId = 0`が残っていた。`Counter`を宣言せずに使う関数は2を返し、既定のメンバーも働いた。全置換は18〜63 msだった。

### `Application.Run`の実行先と受け渡し（[マクロの実行](../architecture.md#マクロの実行)）

- 主ブックの名前で修飾した呼出し: 支持。整数の引数と戻り値（21から42）、日本語の文字列の引数と戻り値、戻り値の無いマクロ（セルに書いた値を読めた）が通った。
- 開いていないブックで修飾した名前（`'NoSuch.xlsm'!Calc.Twice`）: Excelは、Excelの既定のフォルダ（ユーザーのドキュメントのフォルダ）からそのブックを開こうとし、見つからずに`0x800A03EC`で失敗した。ブックは増えなかった。修飾した名前を利用者から受け取ると、アプリのフォルダの外のブックを開いて実行しうる。

### 捕まえない実行時エラーとコンパイルエラー（[マクロの実行](../architecture.md#マクロの実行)）

どちらも、COMの例外としては返らず、ダイアログで止まった（それぞれ3回）。見張りが終了させると、呼出しは`0x800706BE`（`RPC_S_CALL_FAILED`）で戻った。

- 実行時エラー（0で割る）: `#32770`の「Microsoft Visual Basic」のダイアログが表示された。
- コンパイルエラー（宣言の無い変数）: `#32770`の「Microsoft Visual Basic for Applications」のダイアログと、VBEのウィンドウ（`wndclass_desked_gsk`、無効）が表示された。

専用のExcelは非表示でも、このダイアログとVBEのウィンドウは、人の画面に表示された。

### 起動のイベントと`AutomationSecurity`（[起動時の設定](../architecture.md#起動時の設定)）

- `EnableEvents=False`で開いた時: 支持。`Workbook_Open`は動かなかった（`D1`は空）。`EnableEvents=True`で開くと動いた（`D1`は`"opened"`）。
- `AutomationSecurity=Low`で開いた主ブックのマクロの実行: 支持。`VBAWarnings`が2のままで実行できた。
- COMで起動した直後の`AutomationSecurity`は、既に1（`Low`）だった。

### マクロの実行中の終了（[状態と結果不明](../architecture.md#状態と結果不明)）

支持。終わらないマクロ（`DoEvents`の無いループ）の実行中に終了させると、呼出しは約30 ms後に`0x800706BE`で戻った（3回）。この時、表示されたウィンドウは無かった。

## 後始末とGit差分

- 3回の実行の後、どれもEXCELのプロセスは0だった。実行の前からあったプロセスは無かった。終了させたのは、見張りが、所有を確かめた時のハンドルで終了させたExcelだけです。
- `.xlwright/spikes/vba/App/`は、各実行の最後に消した。
- Gitで管理するブックには触れていない。
- Gitの差分は、実験の試行の追加と、本書の追加と、設計とTODO.mdの更新だけです。
