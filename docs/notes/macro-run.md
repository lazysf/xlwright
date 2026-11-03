# マクロの実行時エラーを捕まえる形

`macro_run`（[API](../api.md#macro_run段階1)）の`MACRO_ERROR`と`MACRO_NOT_FOUND`が依存する、実行時エラーの捕まえ方とマクロの有無の見分け方の記録です。
主ブックの名前で修飾した`Application.Run`の受け渡しと、捕まえないエラーがダイアログで止まることは、[VBAの反映とマクロの実行](vba-run.md)にあります。
確かめた日は2026-11-03です。

## 版

- Windows 11 Pro 10.0.26300、.NET 10.0.14（x64）
- Microsoft 365 Excel 64-bit 16.0.20527.20112
- Excelの設定: `AccessVBOM`は1、`VBAWarnings`は2
- 起動した場所: Claude Codeのデスクトップアプリの会話のシェル（Git Bash）から起動した`dotnet`

## 再現の手順

```powershell
dotnet run --project spikes\macro-run -- nested
dotnet run --project spikes\macro-run -- direct
```

実験は、STAのスレッドから専用のExcelを1つ起動し、所有を確かめます（新しいPID、`EXCEL.EXE`、起動を求めた後の生成）。確かめられなければ、触らずに止まります。
ファイルの無い新しいブックに、試すマクロのモジュール`Calc`と、xlwrightのモジュールの候補`XlWright`を取り込みます。1つの呼出しが20秒を過ぎると、所有を確かめたプロセスを終了させます。

- `nested`: `XlWright.Run`（`On Error`を置いた関数）が、マクロの名前を受け取り、`Application.Run`で呼ぶ。
- `direct`: 呼出しごとに、VBIDEで`Calc`の手続きの宣言の行を読み、マクロを名前で直接呼ぶ関数だけの標準モジュール`XlWrightCall`を`VBComponents.Add`と`AddFromString`で作る。主ブックの名前で修飾した`Application.Run`でその関数を呼び、呼出しの後にモジュールを`Remove`する。引数は`(args(0))`のように括弧で囲んで渡す。戻り値は`XlWright.Result`（`ByVal`の`Variant`を受け取る関数）に渡し、オブジェクトと配列を型の名前に変える。

## 判定

### `Application.Run`で呼んだマクロのエラー（`nested`）

不一致（2回）。呼んだ側の`On Error`は、`Application.Run`で呼んだマクロの実行時エラー（0で割る）を捕まえなかった。マクロはダイアログで止まり、20秒後に終了させると、呼出しは`0x800706BE`で戻った。
そのため、マクロの名前を受け取って`Application.Run`で呼ぶ関数では、`MACRO_ERROR`を作れません。

### 生成したモジュールから直接呼んだマクロのエラー（`direct`）

支持。次の実行時エラーを、生成したモジュールの`On Error`で捕まえ、`Err.Number`・`Err.Description`・`Err.Source`を返せた。

| マクロ | `Err.Number` | `Err.Description` |
| --- | --- | --- |
| 0で割る | 11 | 0 で除算しました。 |
| `Err.Raise vbObjectError + 513` | -2147220991 | マクロが渡した日本語の説明 |
| `Private`の手続きの中の型の不一致 | 13 | 型が一致しません。 |
| `ByVal x As Long`に文字列`"x"`を渡す | 13 | 型が一致しません。 |

自分で`On Error Resume Next`を置いたマクロは、エラーにならずに値を返した。
`Err.Source`は、`Err.Raise`で渡した時はその値、ほかは`VBAProject`だった。

### 引数の受け渡し（`direct`）

支持。整数、日本語の文字列、小数・文字列・真偽値の組、30個の引数、`ByRef x As Long`（括弧で囲んで渡した）、`Optional`の省略、`ParamArray`を受け渡せた。
引数の数が宣言と合わない時は、コンパイルエラーのダイアログで止まった。20秒後に終了させた。

### 戻り値の型（`direct`）

.NETが受け取った型です。

| VBAの戻り値 | .NET |
| --- | --- |
| `Long` | `Int32` |
| `String`（日本語、30,000文字） | `String` |
| `Currency` | `Decimal` |
| `Date` | `DateTime` |
| `Empty`と、戻り値の無い`Sub` | `null` |
| `Null` | `DBNull` |
| `CVErr(2042)` | `Int32`（-2146826246） |
| 配列、`Collection`、`Application` | `XlWright.Result`が返した型の名前（`Variant()`・`Collection`・`Application`） |

`CVErr`の値と`Long`は、どちらも`Int32`で届いた。そのため、.NETの型では見分けられず、VBAの側で`VarType`を見る必要があります。
`Application`のように既定のメンバーを持つオブジェクトも、`ByVal`の`Variant`に渡すとオブジェクトのまま渡り、型の名前を返せた。

### マクロの有無（VBIDE）

支持。実行の前に、`VBComponents.Item(<モジュール>)`と`CodeModule.ProcBodyLine(<手続き>, 0)`で見分けられた。

- 無いモジュール: `Item`が`IndexOutOfRangeException`になった。
- 無い手続き: `ProcBodyLine`が`COMException`（`0x800A0023`、「Sub または Function が定義されていません」）になった。
- `CodeModule.Lines(<ProcBodyLine>, 1)`で、宣言の行（`Public Function …`・`Private Function …`・`Public Sub …`）を読めた。行の継続（行末の`_`）がある時は、最初の行だけが返った。

### `End`文

`End`を実行したマクロでは、`Application.Run`が`COMException`（`0x800A9C68`）になった。生成したモジュールは、その後に`Remove`できた。次の呼出しも通った。

### 生成したモジュールの時間

`Add`と`AddFromString`は23〜34 ms、`Remove`は8〜14 msだった。

## 後始末とGit差分

- 2つの試行のどちらも、実行の前と後で、EXCELのプロセスは0だった。ダイアログで止まったExcelは、所有を確かめたプロセスを終了させた。
- 一時フォルダは、実行の最後に消した。
- `nested`は2回、`direct`は1回実行した。どの回も、ダイアログで止まったのは1回で、終了させるまでの20秒、人の画面に表示された見込みです。
- Gitで管理するブックには触れていない。
- Gitの差分は、実験の追加と本書の追加だけです。
