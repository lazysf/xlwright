---
name: xlwright
description: Windows 版 Excel を PowerShell COM で動かし、Excel を土台にしたツール（マクロ、UserForm、カスタムリボンを持ち、人が操作して使うブック）を作って確かめる。VBA と UserForm の作成・編集と入出力、カスタムリボンの埋め込みと見え方・ボタンの動作の確認、ツールのシートの編集、VBA をロックした配布ブックの出力に使う。受け取ったブックの分析、集計の確認、データの整形や変換には使わない。
---

# xlwright

PowerShell から `Excel.Application` を COM で操作して、Excel を土台にしたツールのブックを作り、直す。リボンだけは COM で定義できないため、閉じたブックの zip を直接書き換える。

このスキルのフォルダは、この `SKILL.md` があるフォルダで、スクリプトはその `scripts/` にある。Claude Code では、このフォルダは `${CLAUDE_SKILL_DIR}` である。コマンド例の `<skill>` は、このフォルダの絶対パスに置き換えて実行する。Git Bash からは `powershell.exe -NoProfile -File "<script>"` で実行し、引数のパスは `/` 区切りのまま渡してよい。

## 利用の前提

- Windows 版 Excel（デスクトップ版）、Excel の COM、Windows PowerShell 5.1 を使う。
- WMI（`Win32_Process.Create`）からプロセスを作れること。Excel を起動するたびに、上限で Excel を止める見張りのプロセスを WMI で立てる。立てられない環境では、このスキルは使えない。`Could not start the watchdog` のエラーが出たら、迂回せずに止まり、人へ伝える。
- スクリプトの実行が、実行ポリシー、sandbox、権限で拒否されたら、Excel を使わない手段（ブックの zip や XML を直に組み立てる、ほかのスキルを使う）で代わりをしない。Excel が開いて計算したことを確かめられないからだ。実行環境に権限を引き上げる仕組み（sandbox の外での実行の承認など）があれば、それを求める。無いか、通らなければ、止まって人に伝える。
- エージェントの作業中、人はその端末で Excel を使わない。Excel 以外のアプリは使ってよい。
- リボンの確認（`Test-RibbonLayout.ps1`、`Invoke-RibbonButton.ps1`）と配布ブックの出力では、Excel や VBE のウィンドウが画面に出る。前面は取らず、人が使っているアプリの後ろに出る。出たウィンドウを人が操作しないよう、始める前に伝える。
- VBA のコンパイルの検査（`Test-VbaCompile.ps1`）は、コンパイルエラーがある時だけ、VBE のウィンドウとエラーのダイアログが画面に短い間出る。スクリプトが閉じるので、人は操作しない。
- マクロが出すダイアログ（実行時エラー、コンパイルエラー、MsgBox）は、Excel が非表示でも画面に出る。コンパイルエラーでは VBE のウィンドウも出る。`Invoke-ExcelMacro` が閉じるので、人は操作しない。
- 利用する repo に、ブックの置き場や VBA ファイルの規則（`AGENTS.md`、`CLAUDE.md` など）があれば、そちらを優先する。無ければ、このスキルの既定に従う。

## 作業前の確認

1. **Excel のプロセス**: `Get-Process EXCEL`。動いているなら、ユーザーが開いているブックがある可能性がある。勝手に終了させず、閉じてよいか聞く。
2. **VBA を触るなら `AccessVBOM`**: `HKCU:\Software\Microsoft\Office\16.0\Excel\Security` の `AccessVBOM` が `1` でないと、`VBProject` へのアクセスが失敗する。これはセキュリティ設定なので、自分で変更せず、ユーザーに有効化を頼む。ユーザーに頼まれた時も、確認を取った後も、エージェントは AccessVBOM などのトラストセンターの設定を変えない。手順を渡して、ユーザーに変えてもらう。
3. **バックアップ**: 編集前のブックを scratchpad などの一時領域へコピーする。段階ごと（COM 編集後、customUI 埋め込み前など）にも取ると、どこで壊れたか切り分けられる。
4. **Git 管理**: ブックが Git で追跡されているか、未 commit の変更があるかを確かめる。未 commit の編集や追跡されないブックは、Git では戻せないのでバックアップから戻す。

1、2、4 と、ブックがほかで開かれていないかは、`Test-Prerequisite.ps1` でまとめて調べられる。Excel を起動せず、設定もブックも変えない。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Test-Prerequisite.ps1" -OutFile <result.json> -Path <workbook> -Vba
```

- 結果は UTF-8 の JSON で `-OutFile` に書く。`Ready` が `true` なら進めてよい。`false` の時は、`Problems` に理由と次にすることが 1 件 1 文で入る（Excel が動いている、WMI からプロセスを作れない、ブックが無い、扱わない形式のブック、ブックがほかで開かれている）。終了コードは、`Ready` が `true` の時だけ 0。
- `-Vba` は、VBA を読む・変える作業の時に付ける。付けた時だけ、`AccessVBOM` が `1` でないことを問題にする。
- `Workbook.Git` の `Tracked` と `Uncommitted` は、問題に数えない。追跡されていないブックと、未 commit の変更があるブックは、Git では戻せないので、バックアップを先に取る。
- バックアップ（3）は調べない。ブックを作る前など、ブックがまだ無い時は `-Path` を付けない。
- 結果の項目と、確かめた範囲は、[実行環境と起動の失敗](references/troubleshooting.md)の「前提の確認の結果」にある。

初めての端末や、Excel が更新された後は、[実行環境と起動の失敗](references/troubleshooting.md)に従い、先に「動作確認」の試験を流す。

## マクロが動くことへの注意

Excel は既定で、VBA とブックイベントを無効にし、外部ブックリンクを更新せずに開く。この設定は Excel 4.0 マクロやデータ接続の更新全般を抑止するものではない。

- VBA の動作検証に限り `-EnableMacros` を付ける。実行前に `onLoad` などのコールバックも確認する。`Workbook_Open` などのイベントを検証する時は `-EnableEvents` も付ける。
- マクロの実行（`Invoke-ExcelMacro`、`$xl.Run`）やボタンの押下で動かしてよいのは、この作業で自分が書いた、外部とやり取りしないマクロだけ。既存のマクロやボタンは、ユーザーの確認なしに動かさない。サーバーやファイルを読み書きするマクロは、ブックを保存せずに閉じても、その変更は戻らない。
- sandbox を、マクロに対する防御として数えない。COM で起動した Excel が sandbox の制限の内側で動くとは限らない。

## 手段の分担と順序

| 対象 | 手段 | 理由 |
| --- | --- | --- |
| シート（値、数式、書式、シート構成） | COM（`Invoke-ExcelScript.ps1`） | Excel が再計算し、保存形式も Excel が保つ |
| VBA モジュール、UserForm | COM（`Export-Vba.ps1` / `Import-Vba.ps1`） | `vbaProject.bin` の生成を Excel に任せられる |
| リボン（customUI） | zip を直接編集（`Add-CustomUI.ps1`） | COM にリボンを定義する API が無い |

順序は **COM の編集 → ブックを閉じる → customUI の埋め込み → COM で開き直して検証**。openpyxl など Excel 以外のツールで保存しない（リボンの関連付けが消える。[リボン](references/ribbon.md)）。

## 扱うブックの形式

扱うのは次の 4 つだけ。ほかの形式（`.xltm`、`.xls` など）は、どのスクリプトもブックに触れる前に `Unsupported workbook format` で拒否する。拒否されたら、迂回せずに人へ伝える。

| 形式 | 用途 | 作る（`-Create`） | 開く・編集・確かめる | 配布の作業ブック | 配布の出力 |
| --- | --- | --- | --- | --- | --- |
| `.xlsm` | ツールの本体 | ○ | ○ | ○ | ○ |
| `.xlsb` | ツールの本体（バイナリ形式） | ○ | ○ | ○ | ○ |
| `.xlsx` | ツールが読み書きするブック、ひな形、シートだけのツール | ○ | ○ | × | × |
| `.xlam` | アドインとして配る成果物 | × | ○ | × | ○ |

`.xlsx` は VBA を持てないので、VBA を扱うスクリプトは失敗する。

## 対象と作業手順

ブックごとにフォルダを分け、ブックの隣にソースを置く配置を既定とする。各スクリプトの `-Path` には、そのフォルダのブックを渡す。対象が特定できない場合は、作業ブックの場所を確認してから開く。

```text
<ブックのフォルダ>/
├── <ブック>.xlsm
├── vba/
│   ├── standard/
│   ├── classes/
│   ├── documents/
│   ├── forms/
│   └── references.json
└── ribbon/
    └── customUI14.xml
```

| 作業 | 読むもの |
| --- | --- |
| 手を入れる前に、既存のツールのブックに何があるかを調べる（シート、定義名、テーブル、外部リンク、データ接続、VBA、リボン） | `Get-WorkbookSummary.ps1`。[COM 編集の詳細](references/com.md)の「ブックの概要」 |
| ツールのシート（入力欄、設定、数式）の確認・編集 | 下の「COM 編集」。確認だけなら `-Save` を付けない |
| ブックが人にどう見えるかを画像で確かめる（書式、グラフ、ボタン、リボンを含む画面） | `Save-WorkbookScreenshot.ps1`。[COM 編集の詳細](references/com.md)の「見た目を画像にする」 |
| VBA の変更（`vba/` の正本を編集して取り込む） | [VBA の編集ルールと入出力](references/vba.md) |
| UserForm の作成・変更 | [UserForm の作成と編集](references/userform.md) |
| リボンの変更、見え方とボタンの動作確認 | [リボン](references/ribbon.md) |
| VBA をロックした配布ブック（`.xlsm`、`.xlsb`）かアドイン（`.xlam`）の出力 | [配布ブックの出力](references/publish.md) |
| ツールの作業で複数のブックを順に扱う、ひな形のシートを別ブックからコピーする | [COM 編集の詳細](references/com.md) |
| Excel が起動しない、実行が拒否される | [実行環境と起動の失敗](references/troubleshooting.md) |

編集後はブックを開き直し、「検証」で変更した箇所を確認する。

ブックから VBA のコンポーネントを消すのは、正本からモジュールを消した時だけにする。人がブックに足したコードも、正本に無いものとして差に出る。消す前に、書き出しと正本の差で「ブックにだけある」コンポーネントを調べ、名前をユーザーに示して確認を取り、確認した名前だけを消す（[VBA の編集ルールと入出力](references/vba.md)の「正本から消したモジュール」）。

### 新しいブックを作る時の流れ

VBA とリボンを持つブックを一から作る時は、次の順に進める。各手順の詳細は上の表の reference にある。

1. `Invoke-ExcelScript.ps1 -Create -Save` でブックを作り、シートの値・数式・書式を入れる。
2. `vba/` に VBA のソースを書き、`Import-Vba.ps1` で取り込む。取り込んだら `Test-VbaCompile.ps1` でコンパイルを確かめる。直しては確かめる繰り返しでは、取り込み、コンパイル、マクロの試験、読み戻しを 1 回の呼び出しにまとめる（「検証」の「まとめて確かめる」）。
3. ブックを閉じた状態で、`ribbon/customUI14.xml` を `Add-CustomUI.ps1` で埋め込む。
4. `Invoke-ExcelScript.ps1` で開き直し、シートと VBA を読み戻す。リボンは `Test-RibbonLayout.ps1` で見え方を XML と照らし、`Invoke-RibbonButton.ps1` で押して確かめる。
5. 配布する場合は、最後に `Publish-Workbook.ps1` で出力し、ロックを読み戻す。

## COM 編集

Excel を `New-Object -ComObject` で直接起動しない。起動、ブックを開く、後始末は `XlWright.psm1` の `Use-ExcelWorkbook` に任せる。使い捨ての操作は、ScriptBlock の中身だけをファイルに書き、`Invoke-ExcelScript.ps1` に渡す。ファイルは scratchpad などの一時領域に置く。

```powershell
# read.ps1（UTF-8。日本語を書いてよい）
param($xl, $wb)
$ws = $wb.Worksheets.Item('集計表')
@{ Rows = (Read-ExcelRange -Range $ws.Range('A1:B2')); Formulas = (Read-ExcelRange -Range $ws.Range('A1:B2') -Formula) }
```

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Invoke-ExcelScript.ps1" -Path <workbook> -Script <read.ps1> -OutFile <result.json>
```

- スクリプトのファイルは UTF-8 として読む（BOM の有無を問わない）。日本語のシート名や文字列を直接書いてよい。
- 結果は UTF-8 の JSON で `-OutFile` に書く。`Ok`、`Result`（ScriptBlock が出力した値の配列）、`Error`、`Warnings` を持つ。失敗した時は終了コードが 1。`powershell.exe` の標準出力は Git Bash から見ると日本語が化けるので、結果は JSON で読む。
- 返す値は文字列・数値・真偽値・hashtable・配列にする。COM オブジェクトは `"[COM object]"` に置き換わる。複数セルの `Value2`（2 次元配列）は行の配列になるが、エラー値は数値（`-2146826281` など）のまま出るので、セルは `Read-ExcelRange` で読む。COM メソッドの戻り値（`Worksheets.Add()` など）も結果に混ざるので、要らないものは `[void]` で捨てる。
- `-Create` を付けると、先に新しいブックを作る（既にあれば失敗する）。拡張子が `.xlsm` ならマクロ有効ブック、`.xlsb` ならバイナリ形式のブック、`.xlsx` なら通常のブックになる。作れる形式は「扱うブックの形式」のとおり。
- `-Save` は、ScriptBlock が例外なしで終わった時だけ保存する。例外が出たら、保存せずに閉じる。
- `-Visible` はウィンドウを表示する。ユーザーの画面に Excel が出るので、先に一言伝える。
- `-TimeoutSeconds`（既定 300）を過ぎると、見張りのプロセスがこの呼び出しで起動した Excel だけを止め、`Timed out after ...` のエラーになる。保存していない変更は失われる。見張りは呼び出し元のシェルが timeout などで殺されても残り、上限で Excel を止める。長いマクロを動かす時は、上限を明示して延ばす。
- ブックを閉じて `Quit` した後、Excel が 5 秒以内に終わらなければ、この呼び出しが起動した Excel（PID と起動時刻で照合）を止める。PowerShell が COM の参照を手放さず、`Quit` の後も Excel が数分残る実測があるため。
- 警告で PID が出た時（照合できない、止められない）は、人へ伝える。自分で止める時も、その PID だけを `Stop-Process -Id` で止める。名前（`-Name EXCEL`）では止めない。ユーザーの Excel まで止まる。

### マクロを動かす

マクロは、ScriptBlock の中で `Invoke-ExcelMacro` を使って動かす。`-EnableMacros` を付けて開く。

```powershell
param($xl, $wb)
Invoke-ExcelMacro -Excel $xl -Name 'Util.AddNums' -ArgumentList 2, 3
```

`$xl.Run` は、マクロがダイアログを出している間は戻らず、非表示の Excel では上限まで止まったままになる。`Invoke-ExcelMacro` は、マクロの実行中に出たダイアログを見張って、次のように扱う。

- **実行時エラー**: ダイアログを「終了」で閉じ、その文言（「実行時エラー '11': 0 で除算しました。」など）を入れた例外にする。Excel は止まらず、続けて使える。
- **コンパイルエラー**: ダイアログを閉じ、VBE をリセットして隠し、その文言（「コンパイル エラー: 変数が定義されていません。」など）を入れた例外にする。Excel は止まらず、続けて使える。VBE のウィンドウが短い間画面に出る。エラーの場所は分からないので、`Test-VbaCompile.ps1` で調べる。
- **OK だけの MsgBox**: 閉じて先へ進む。題名と本文を警告に出すので、`Invoke-ExcelScript.ps1` の結果の `Warnings` で読める。
- **閉じられないダイアログ**（はい/いいえ、OK/キャンセル の MsgBox など、答えを選ぶもの）: 文言を読んでから、この呼び出しが起動した Excel を止め、文言を入れた例外にする。保存していない変更は失われ、`Workbook.Close failed` などの警告も出る。答えを選ぶ MsgBox を出すマクロは、確かめたい分岐を引数で選べる形に分けてから動かす。コンパイルエラーの後に VBE をリセットできなかった時も、同じ扱いになる。

戻り値はマクロの戻り値で、`Sub` は何も返さない。引数は 30 個まで。名前はモジュール名で修飾する。同じ名前の `Sub` がほかのモジュールにあると、修飾しない呼び出しは失敗する。実行時エラーの後は、モジュール変数が初期値に戻るので、前の呼び出しが残した状態に頼らない。InputBox、ファイルを選ぶダイアログ、UserForm を出すマクロでは確かめていない。

### PowerShell + COM の癖

- **複数のセルは `Read-ExcelRange` と `Write-ExcelRange` で読み書きする。** 1 セルずつの代入は遅く、同じ行の代入に型の違う値を渡した時や、`ConvertFrom-Json` の数値を代入した後の代入が `InvalidCastException` で落ちる。一度落ちると、同じプロセスでは配列の代入も落ち続ける。2 つの関数はこれを避け、`001` や `1/2` のような文字列も文字のまま書く。使い方は [COM 編集の詳細](references/com.md)の「セルの範囲を読み書きする」。
- **1 セルに書く時も、`Write-ExcelRange -Range $ws.Range('B2') -Value (, @($v))` を使う。** `$ws.Range('B2').Value2 = $v` の形は、上の `InvalidCastException` に当たることがあり、文字列は Excel が解釈して変える。
- **`-File` で直接実行するスクリプトは ASCII だけで書く。** `powershell.exe`（5.1）は BOM なしの `.ps1` を UTF-8 として読まない。`Invoke-ExcelScript.ps1` に渡すファイルには、この制限はない。

## 検証

編集した後は、`-Save` なしでブックを開き直して読み戻す。書いたつもりの値ではなく、Excel が保存した結果を見る。動かしてよいマクロの範囲は「マクロが動くことへの注意」のとおり。

- **シート**: `Value2`、`Formula`、`Text`（表示形式が適用された文字列）を読む。数式は Excel が再計算した値で確かめる。数式を書いた・行や列やシートを消した後は、`Find-CellError.ps1` で `#REF!` や `#NAME?` のセルが無いかを調べる（[COM 編集の詳細](references/com.md)の「エラー値のセルを探す」）。書式、条件付き書式、列幅、グラフ、ボタンの見た目は、`Save-WorkbookScreenshot.ps1` で画像にして見る（同じく「見た目を画像にする」）。
- **VBA**: 変更したコンポーネントの名前・種別・コードを読み戻す。`Test-VbaCompile.ps1` でコンパイルし、構文エラーや宣言の無い変数が無いことを確かめる（[VBA の編集ルールと入出力](references/vba.md)）。実行してよいマクロは `Invoke-ExcelMacro` で動かし、戻り値を確かめる（「マクロを動かす」）。確かめる場合が多い時も、ScriptBlock の中で `Invoke-ExcelMacro` を並べて呼ぶ。確かめるためだけの VBA は、`vba/` にも作業ブックにも残さない。`Publish-Workbook.ps1` は作業ブックの写しに `vba/` を取り込み、取り込みは正本に無いコンポーネントを消さないので、どちらに残したコードも配布ブックに入る。要る時は、一時フォルダに取ったブックの写しに足す。ThisWorkbook に置いた関数は、`$xl.Run` で呼ぶと（ブック名で修飾しても）空が返る。`[System.__ComObject].InvokeMember('FuncName', [Reflection.BindingFlags]::InvokeMethod, $null, $wb, $null)` で Workbook のメンバーとして呼ぶと値が返る。
- **リボンとボタン**: `Test-RibbonLayout.ps1` でタブ・グループ・ボタンの並び、大きさ、有効無効、表示非表示を XML と照らし、タブの画像でアイコンを見る。`Invoke-RibbonButton.ps1` で実際に押し、出たダイアログの文言を確かめる（[リボン](references/ribbon.md)）。

### まとめて確かめる

スクリプトは呼ぶたびに Excel を起動して閉じ、測った端末では 1 回に約 4 秒かかった。ブックを開いて作業する時間は、そのうち 0.5 秒ほど。取り込み、コンパイル、マクロの試験、読み戻し、エラー値のセルの検査は、`Invoke-ExcelScript.ps1` の 1 回の呼び出しにまとめる。ScriptBlock の中では、各スクリプトと同じ処理を関数で呼べる。

```powershell
# check.ps1（UTF-8）。ブックの隣の vba/ を取り込み、同じ Excel で確かめる
param($xl, $wb)
$imported = @(Import-ExcelVba -Workbook $wb -SourceDir (Join-Path $wb.Path 'vba'))
$compile = Test-ExcelVbaCompile -Excel $xl -Workbook $wb
$result = [ordered]@{ Imported = $imported; Compile = $compile }
if ($compile.Compiled) {
    $result.Sum = Invoke-ExcelMacro -Excel $xl -Name 'Util.AddNums' -ArgumentList 2, 3
    $result.Cells = Read-ExcelRange -Range $wb.Worksheets.Item(1).Range('A1:B2')
    $result.Errors = Find-ExcelCellError -Excel $xl -Workbook $wb
}
$result
```

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Invoke-ExcelScript.ps1" -Path <workbook> -Script <check.ps1> -OutFile <result.json> -EnableMacros -Save
```

| 関数 | 同じ処理のスクリプト | 返すもの |
| --- | --- | --- |
| `Import-ExcelVba -Workbook $wb -SourceDir <vba>` | `Import-Vba.ps1` | `imported: <名前>` などの行。保存はしない |
| `Test-ExcelVbaCompile -Excel $xl -Workbook $wb` | `Test-VbaCompile.ps1` | `Compiled` と `CompileError` |
| `Find-ExcelCellError -Excel $xl -Workbook $wb` | `Find-CellError.ps1` | `Count`、`Truncated`、`Calculation`、`Sheets` |

- 同じ呼び出しで読むのは、保存する前の Excel の中の状態。直しては確かめる繰り返しの間はこれで足りる。編集を終えたら、上のとおり `-Save` なしで開き直して、保存された結果を読み戻す。
- 取り込みを残すには `-Save` を付ける。コンパイルエラーは例外にならず `Compiled` が `false` で返るので、`-Save` を付けていれば取り込んだコードは保存される。別々のスクリプトで取り込んでからコンパイルした時と同じ。
- コンパイルエラーの時は、VBE とダイアログが画面に短い間出る。閉じた後も、同じ Excel でエラーの無いモジュールのマクロは動いた。エラーのあるモジュールのマクロを動かした時は確かめていない。
- `-EnableMacros` は、マクロを動かす時だけ付ける。動かしてよいマクロの範囲は「マクロが動くことへの注意」のとおり。
- 閉じたブックが要る手順（`Add-CustomUI.ps1`）と、Excel を表示する手順（リボンの確認、`Save-WorkbookScreenshot.ps1`、配布）は、まとめられない。別の呼び出しにする。
- 引数と結果の項目は、同じ処理のスクリプトの説明（[VBA の編集ルールと入出力](references/vba.md)、[COM 編集の詳細](references/com.md)の「エラー値のセルを探す」）のとおり。

## 動作確認

スクリプトが手元の Excel で動くかは、`tests/Test-XlWright.ps1` で確かめる。各項目は一時フォルダに新しいブックを作って、期待値と照合する。一時フォルダの外は変えない。Excel が動いている時と `AccessVBOM` が `1` でない時は、何もせずに終了コード 2 で止まる。

```powershell
powershell.exe -NoProfile -File "<skill>/tests/Test-XlWright.ps1" -OutFile <result.json>
```

- 既定の項目は Excel を表示しない。セルと VBA、`Invoke-ExcelScript.ps1`、範囲の読み書き、VBA のコンパイル、マクロの実行、上限、呼び出し元が殺された時、不正な入力の拒否（VBA のソース、リボンの XML、配布の失敗）、見張りが立てられない時、前提の確認、ブックの概要、エラー値のセルの列挙、1 回の呼び出しでまとめた確認、UserForm を確かめる。5 分ほどかかる。
- `-Full` を付けると、コンパイルエラーの検出、マクロが出すダイアログの扱い、リボンの押下（同名の組み合わせ、閉じられないダイアログを含む）、リボンの見え方の照合、配布ブックの出力、ブックと UserForm の見た目の画像も確かめる。Excel と VBE が画面に出るので、先に人へ伝え、出たウィンドウを操作しないよう頼む。これらの項目は、Excel が前面を取らないことも確かめる。
- `-Case timeout,userform` のように、名前を指定した項目だけを流せる。
- 終了コードは、0 がすべて合格、1 が不合格あり、2 が前提を満たさない・Excel が残って止まった。一時フォルダは合否によらず最後に消す。不合格の理由は出力と `-OutFile` に残る。試験のブックを調べたい時は、`-KeepTemp` を付けて流し直すと残り、場所を出力する。
