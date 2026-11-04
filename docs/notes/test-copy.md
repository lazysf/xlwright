# テストの写しとコンパイルとシートのモジュール

段階3の`test_run`・`vba_check`・シートとブックのモジュールの反映が依存する、設計の未確認の5つの記録です（[テスト](../architecture.md#テスト)、[VBAの反映](../architecture.md#vbaの反映)）。
確かめた日は2026-11-04です。

## 版

- Windows 11 Pro 10.0.26300、.NET 10.0.14（x64）
- Microsoft 365 Excel 64-bit 16.0.20527.20112
- Excelの設定: `AccessVBOM`は1、`VBAWarnings`は2
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`dotnet`

## 再現の手順

```powershell
dotnet run --project spikes\test-copy -- flow
dotnet run --project spikes\test-copy -- compile close
dotnet run --project spikes\test-copy -- compile ok
dotnet run --project spikes\test-copy -- compile keep
dotnet run --project spikes\test-copy -- document
```

実験は、STAのスレッドから専用のExcelを1つ起動し、所有を確かめます（新しいPID、`EXCEL.EXE`、起動を求めた後の生成）。確かめられなければ、触らずに止まります。
ブックは、実行ごとに作る一時フォルダにだけ作ります。専用のExcelのダイアログを200ミリ秒ごとに探して文字を表示し、2つ目の引数に従って扱います（`close`は`WM_CLOSE`、`ok`は`IDOK`の`WM_COMMAND`、`keep`は残す）。1つの手順が30秒で戻らなければ、所有を確かめたプロセスを終了させます。

- `flow`: `Workbook_Open`と`Auto_Open`が印のファイルを作る主ブック`Main.xlsm`を、`data/`・`src/`・`tests/`・`.git/`のあるアプリのフォルダに作る。設計の流れの1〜5を行う。写すのは`src/`・`tests/`と`.`で始まるものを除くすべて。`tests/Tests.bas`は、`ThisWorkbook.FullName`と`Workbooks("Main.xlsm").FullName`を返す関数と、シートと`data/note.txt`を書き換える手続き。開く時は`AutomationSecurity=Low`と`EnableEvents=False`、テストの実行中は`EnableEvents=True`。最後に、対照として主ブックを`EnableEvents=True`で開く。
- `compile`: 新しいブックの`VBE.CommandBars.FindControl(1, 578)`（「VBAProject のコンパイル」）を、`ActiveVBProject`をそのブックにして`Execute`する。正しいコード、`Option Explicit`での宣言の無い変数、構文エラー、エラーのモジュールを消した後、2つ目のブックが別のエラーを持つ時の順。
- `document`: 新しいブックのシートの名前を変え、`VBComponents.Item(<CodeName>)`のシートとブックのモジュールのコードを`DeleteLines`と`AddFromString`で書き換える。保存して開き直し、`Worksheet_Change`を起こす。

## 判定

### テストの写しの流れ

支持。保存して閉じ、写し、写しを開いて`tests/Tests.bas`を`Import`し、実行し、保存せずに閉じて写しのフォルダを消し、主ブックを開き直せた。
写しのシートと`data/note.txt`はテストで書き換わり、主ブックのセルと`data/note.txt`は元のままだった。開き直した主ブックに`Tests`は無かった。

### 起動のイベントの抑止

支持。`EnableEvents=False`で開くと、主ブックでも写しでも`Workbook_Open`は動かなかった。対照の`EnableEvents=True`では動いた。
`Auto_Open`は、どちらでも動かなかった。外部のCOMの`Workbooks.Open`では、`Auto_Open`は動かない。

### 実行先を写しに限れるか

支持。主ブックを閉じた後は、主ブックの名前で修飾した`'Main.xlsm'!Tests.TestWhere`が写しのマクロを呼んだ。`ThisWorkbook.FullName`も`Workbooks("Main.xlsm").FullName`も、写しのパスを返した。
写しを開いている間に、`DisplayAlerts=False`で主ブックを`Workbooks.Open`すると、例外もダイアログも無く`null`が返り、開いたブックは写しの1つのままだった。同じ名前のブックは同時に開けないため、主ブックの名前は写しだけを指す。
テストが主ブックのフォルダを絶対パスで書き換えることは、防げない。

### 外部のCOMからのコンパイル

支持。`Execute`でコンパイルが起き、正しいコードではダイアログが出ず、コントロールの`Enabled`が偽になった。モジュールを加えると真に戻り、エラーのモジュールを消すと偽に戻った。`Enabled`が偽の時の`Execute`は`0x80004005`で失敗した。
`ActiveVBProject`のブックだけがコンパイルされた。2つ目のブックにエラーがあっても、1つ目を`ActiveVBProject`にした`Execute`はダイアログを出さなかった。

### 非表示のExcelで、コンパイルのエラーがダイアログで止まらないか

不一致。呼出しは止まらないが、ダイアログは出る。

- `Execute`は0ミリ秒で戻り、その後2秒以内に「Microsoft Visual Basic for Applications」のダイアログが出た。本文は`コンパイル エラー:`に続く説明（「変数が定義されていません。」「構文エラー」）。
- ダイアログと同時に、VBEのウィンドウが人の画面に表示された（`VBE.MainWindow.Visible`が真）。Excelの`Visible`は偽のままだった。`MainWindow.Visible=False`で隠せた。
- コントロールの`Enabled`は真のままだった。`ActiveCodePane.GetSelection`を表示した3回では、宣言の無い変数の時はエラーの行（4行目の5〜18列）を毎回指した。構文エラーの時は、5行目を1回指し、2回は`ActiveCodePane`が無かった。
- ダイアログを開いたままでも、以後のCOMの呼出しは通った。
- ダイアログを`WM_CLOSE`で閉じた3回は、Quitで終わった。`IDOK`で閉じた1回と残した1回は、Quitの後20秒で終わらず、終了させた。
- 最初の1回は、変数の名前がモジュールの名前と同じだったため、本文が「モジュールではなく、変数またはプロシージャを指定してください。」になった。この回は、`WM_CLOSE`の後にExcelが終わった（`0x800706BA`）。原因は確かめていない。

### シートとブックのモジュールの書き換え

支持。`VBComponents.Item(<CodeName>)`で種類が100（ドキュメント）のモジュールを見つけ、`DeleteLines`と`AddFromString`で書き換えられた。2回書いても追記にならず、日本語を含むコードを読み戻すと一致した。モジュールの数は変わらなかった。
保存して開き直した後もコードは残り、`EnableEvents=True`でセルを書くと`Worksheet_Change`が日本語の値を書いた。
シートの名前を「データ」に変えても、`CodeName`は変わらなかった。無い名前の`Item`は例外になった（`0x80131508`）。

`Workbooks.Add`で作ったばかりのブックでは、`VBProject`に触れる前の`CodeName`はブックもシートも空だった。触れた後に、シートの並びの順で付き、元の`Sheet1`のシートに`Sheet2`が付いた。保存して開き直した後は、同じ`CodeName`のままだった。

## 未確認

- `AddFromString`に、システムのANSIコードページに無い文字を渡した時の扱い。
- ブックのモジュールの`Public Function`を`Application.Run "'Doc.xlsm'!ThisWorkbook.Hello"`で呼ぶと、戻り値が空の文字列だった。`macro_run`は標準モジュールに限るため、追っていない。
- `compile close`の1回目は、起動の直後の`Application.Visible`が真だった。ほかの回は偽。原因は確かめていない。

## 後始末とGit差分

- どの回も、前後の`EXCEL.EXE`は0で、所有を確かめたプロセスだけを終えた。
- 一時フォルダは、どの回も消えた。
- Gitの差分は、`spikes/test-copy/`と文書だけ。管理するブックに触れていない。
