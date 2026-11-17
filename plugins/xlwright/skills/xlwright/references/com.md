# COM 編集の詳細

コマンド例の `<skill>` は、SKILL.md に書いたこのスキルのフォルダの絶対パスに置き換える。

## ブックの概要

既存のブックに何が入っているかを、編集の前に 1 回で調べる。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Get-WorkbookSummary.ps1" -Path <workbook> -OutFile <result.json>
```

- 非表示の Excel で、マクロとイベントを無効にし、リンクを更新せずに開いて、保存せずに閉じる。ブックは変わらない。
- 結果は UTF-8 の JSON で `-OutFile` に書く。終了コードは、読めた時に 0。

  | 項目 | 中身 |
  | --- | --- |
  | `Sheets` | ブックの中の並びの順。`Name`、`CodeName`、`Kind`（`worksheet`、`chart`、`other`）、`Visible`（`visible`、`hidden`、`veryHidden`）、`UsedRange`、`Rows`、`Columns`、`Protected`、`Tables`（`Name`、`Range`） |
  | `Names` | 定義名。`Name`、`Scope`（`workbook` か、シートの名前）、`RefersTo`、`Visible` |
  | `ExternalLinks` | 数式がリンクしているブックのパス |
  | `Connections` | データ接続。`Name` と `Type`（`oledb`、`odbc`、`xml`、`text`、`web`、`datafeed`、`model`、`worksheet`、`nosource`、`other`） |
  | `Vba` | `HasProject`、`Locked`、`Components`（`Name`、`Kind`、`Lines`）、`Error` |
  | `CustomUI` | zip の `customUI/` にある項目の名前。リボンが無ければ空 |
  | `ProtectStructure` | ブックの構成が保護されているか |

- `UsedRange` は、値も数式も無いシートと、ワークシート以外のシートでは `null` になる（`Rows` と `Columns` は 0）。Excel は空のシートの `UsedRange` を `A1` と返すので、`A1` にだけ値があるシートと見分けるために分けている。書式だけを付けたシートも `null` になる作りだが、確かめていない。
- `Names` には、Excel が自動で作った名前（外部データの `ExternalData_1` など）も入る。
- `Connections` に、接続文字列と SQL は入れない。サーバー名や認証情報が入っていることがあるため。中身が要る時は、ユーザーに確かめてから読む。
- `Vba.Components` の `Kind` は、`standard`、`classes`、`documents`、`forms`、`other`。コードの無い文書モジュールは省く。VBA を読めない時（`AccessVBOM` が `1` でない、VBA プロジェクトがロックされている）は、`Components` が `null` になり、`Vba.Error` に理由が入る。ほかの項目は返り、終了コードは 0 のまま。
- ブックが無い時や開けない時は、`Error` に理由が入り、終了コードは 1 になる。
- 確かめたのは、`.xlsm` と `.xlsx`、グラフシート、非表示のシート、テーブル、CSV を読み込んだ接続、ブックへのリンク 1 つ、VBA プロジェクトをロックしたブック、リボンを埋め込んだ `.xlsb` の `CustomUI`。OLE DB や ODBC の接続、Power Query、リンク先の無いリンク、保護されたシート、Excel 4.0 マクロシート、`AccessVBOM` が `1` でない時、`.xlsb` の `CustomUI` のほかの項目、`.xls` では確かめていない。

## エラー値のセルを探す

編集の後に、値がエラー（`#REF!`、`#NAME?`、`#DIV/0!` など）になったセルを、ブック全体から 1 回で列挙する。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Find-CellError.ps1" -Path <workbook> -OutFile <result.json>
```

- 非表示の Excel で、マクロとイベントを無効にし、リンクを更新せずに開いて、保存せずに閉じる。ブックは変わらない。計算はやり直さない。
- ワークシートの使っている範囲を全部読む。非表示のシートと保護したシートも読む。グラフシートは読まない。
- `-Sheet Data,Summary` のように名前を渡すと、そのワークシートだけを読む。ワークシートでない名前や無い名前は、失敗になる。
- `-MaxErrors`（既定 1000）を超えた分は `Errors` に入れず、数だけを `Count` に足す。
- 結果は UTF-8 の JSON で `-OutFile` に書く。終了コードは、読めた時に 0（エラーのセルがあっても 0）、ブックが無い時や開けない時に 1。

  | 項目 | 中身 |
  | --- | --- |
  | `Count` | 見つかったエラーのセルの数 |
  | `Truncated` | `Errors` が `-MaxErrors` で打ち切られたか |
  | `Calculation` | Excel の計算方法（`automatic`、`manual`、`semiautomatic`、`other`） |
  | `Sheets` | ブックの中の並びの順。`Name`、`Count`、`Errors` |
  | `Errors` の各項目 | `Address`（`B3` の形）、`Error`（`#REF!` など）、`Kind`、`Formula`、`SpillParent` |
  | `Error`、`Warnings` | 失敗の理由と、閉じる時の警告 |

- `Kind` は、数式のセルが `formula`、エラーを打ち込んだセルが `constant`、ほかのセルの数式がこぼした先が `spill`。`Formula` は `formula` のセルだけに `Formula2` が入る。`SpillParent` は `spill` のセルだけに、こぼした数式の番地が入る。
- こぼれる先がふさがった数式は、そのセルが `#SPILL!` の `formula` になる。こぼれた先のエラーは、こぼした数式と別に 1 セルずつ数える。
- `Calculation` が `manual` の時は、保存された値が古いことがある。直したはずのエラーが残って見えたり、新しいエラーが見えなかったりする。その時は、結果が古いかもしれないことをユーザーに伝える。
- 定義名、条件付き書式、入力規則の中の `#REF!` は、セルの値に出ないので見つからない。定義名の参照先は `Get-WorkbookSummary.ps1` の `Names` で読める。
- 確かめたのは、`.xlsx` の数式と定数のエラー、こぼれた先のエラー、`#SPILL!`、非表示と保護のシート、使っている範囲が `A1` から始まらないシート、20001 行のシート。数式を非表示にした保護シート、手動の計算で保存されたブック、16384 列に近い広いシート、`.xlsb` と `.xls` では確かめていない。

## セルの範囲を読み書きする

`XlWright.psm1` の `Read-ExcelRange` と `Write-ExcelRange` は、範囲を 1 回の COM の呼び出しで読み書きする。`Use-ExcelWorkbook` の ScriptBlock の中と、`Invoke-ExcelScript.ps1` に渡すスクリプトの中で使える。

```powershell
param($xl, $wb)
$ws = $wb.Worksheets.Item(1)
Write-ExcelRange -Range $ws.Range('A1') -Value @(@('Code', 'Qty'), @('001', 3), @('002', $null))
Write-ExcelRange -Range $ws.Range('C2') -Formula -Value @(@('=B2*2'), @('=B3*2'))
@{ Values = (Read-ExcelRange -Range $ws.Range('A1:C3')); Formulas = (Read-ExcelRange -Range $ws.Range('C2:C3') -Formula) }
```

### `Read-ExcelRange -Range <Range> [-Formula]`

- 行の配列を返す。`$rows[0][0]` が左上のセル。1 セルや 1 行の範囲でも、同じ形（1 セルなら `@(, @(<値>))`）で返る。`Invoke-ExcelScript.ps1` の結果では、行の配列の JSON になる。
- 値は `Value2` のとおり。空のセルは `null`、文字列、数値（`Double`）、真偽値。日付はシリアル値の数値で、表示形式は付かない。
- エラー値は `@{ Error = '#DIV/0!' }` になる。名前は `#NULL!`、`#DIV/0!`、`#VALUE!`、`#REF!`、`#NAME?`、`#NUM!`、`#N/A`、`#GETTING_DATA`、`#SPILL!`、`#CALC!`。どれでもないコードは `#ERROR<コード>` になる。表示の言語によらない。
- `-Formula` を付けると、`Formula2` の文字列を返す。数式のセルは数式、値だけのセルはその値の文字列、空のセルとこぼれた先のセルは空の文字列。文字として書いた `=A1` も `=A1` で返り、数式と見分けられない。見分ける時は、`-Formula` なしで読んだ値と比べる（数式なら計算した値になる）。
- 離れた範囲（`A1:B2,D1:E2`）は `ArgumentException` になる。`Value2` は例外を出さず、最初の範囲だけを返すため。

### `Write-ExcelRange -Range <Range> -Value <rows> [-Formula]`

- `-Value` は、長さのそろった行の配列。`Read-ExcelRange` の戻り値をそのまま渡せる。1 行は `(, @(1, 2, 3))`、1 列は `@(@(1), @(2))` と書く。
- `-Range` は、左上のセルか、`-Value` とちょうど同じ大きさの範囲。大きさの違う範囲、離れた範囲、シートの端を越える大きさは `ArgumentException` になる。Excel は大きさの違いを例外にせず、`#N/A` で埋めるか、はみ出た分を捨てるため。
- セルに書けるのは、`null`、文字列、数値の型、真偽値、`DateTime`。PowerShell の PSObject で包んだ値は、中の値を書く。
  - `null` と空の文字列は、セルを空にする。
  - `DateTime` はシリアル値で書き、表示形式は変えない。日付として見せる時は、`NumberFormat` を別に設定する。
  - `char`、`Guid`、hashtable、配列、`NaN` と無限大、32767 字を超える文字列は `ArgumentException` になる。Excel に渡すと、例外にならずに文字コード、空のセル、`#NUM!`（`NaN`）、切り詰めた文字列に変わるか、Excel の中で失敗した。無限大を Excel に渡した時は確かめていない。
  - 形と型の誤りは書く前に調べるので、`ArgumentException` の時は何も書かれていない。
- 文字列は文字のまま書く。`001`、`1/2`、`TRUE`、`=A1`、` 12 ` も、そのまま文字列として入る。
  - Excel は配列の文字列を、人が打ち込んだ時と同じように解釈する（`001` は数値の 1、`1/2` は日付）。そこで、一度書いて読み戻し、変わった文字列だけを先頭に `'` を付けて書き直す。そのセルは `PrefixCharacter` が `'` になる。値（`Value2`、`Formula`、`Text`）に `'` は入らない。
  - 先頭が `=` か全角の `＝` の文字列は、Excel が受け取った時点で数式として計算する。そのため、最初の書き込みから `'` を付ける。外から来たデータの `=WEBSERVICE(...)` のような文字列も、計算されない。
  - 書き直しても文字のままにならないセルがあれば、そのセルを挙げて例外にする。範囲は書かれた後になる。
- `-Formula` を付けると `Formula2` に書き、Excel が打ち込んだ時と同じように解釈する。`=` で始まる文字列は数式になり、こぼれる数式（`=SEQUENCE(3)`）は下のセルへこぼれる。`001` は数値の 1 になる。
- 何も返さない。
- 保護されたシートなど、書き始めた後に Excel が失敗した時は、その例外がそのまま出る。範囲は途中まで書かれていることがある。`Use-ExcelWorkbook` と `Invoke-ExcelScript.ps1` の `-Save` は、例外の時は保存しない。
- 100000 行 × 10 列の数値で、一括の書き込みは約 1 秒、読み出しは約 0.3 秒だった。1 セルずつ書くと、100 行 × 10 列で約 3 秒かかった。関数そのものの時間は測っていない。

確かめたのは、日本語の表示の Microsoft 365 Excel。日付の書き方が違う地域の設定、テーブルの中、入力規則のあるセル、結合したセルへの書き込みでは確かめていない。結合したセルは、左上のセルだけが値を持つ。

## モジュールを直接使う

複数のブックを順に扱うなど、1 回の ScriptBlock に収まらない時は、スクリプトから `XlWright.psm1` を読み込む。スクリプトは一時領域に置き、`powershell.exe -NoProfile -File <script> -Path <workbook>` で実行する。

```powershell
param([string]$Path)
$ErrorActionPreference = 'Stop'
Import-Module '<skill>/scripts/XlWright.psm1'

Use-ExcelWorkbook -Path $Path -Save -ScriptBlock {
    param($xl, $wb)
    $ws = $wb.Worksheets.Item(1)
    Write-ExcelRange -Range $ws.Range('A1') -Value @(@('Item', 'Price', 'Count'), @('A-001', 3, 2))
    Write-ExcelRange -Range $ws.Range('D2') -Formula -Value (, @('=B2*C2'))
}
```

- **この形のスクリプトは ASCII だけで書く。** `powershell.exe`（5.1）は BOM なしの `.ps1` を UTF-8 として読まない。日本語のシート名、モジュール名、VBA コードは、UTF-8 のデータファイルに分けて `[IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)` で読み込む。この方法で、日本語のモジュール名、コメント、文字列リテラル（全角と半角カナを含む）が化けずに入り、`$xl.Run` の戻り値も一致することを確かめてある。日本語のパスやラベルを `-File` の引数として渡すのは問題ない（Git Bash から `-TabLabel 'ツール'` のように渡して動くことを確かめてある）。
- **日本語の結果はファイルに書いて読む。** 確認したい値は `[IO.File]::WriteAllLines($p, $lines, (New-Object Text.UTF8Encoding($false)))` で UTF-8 のファイルに書き出す。
- 新しいブックは `New-ExcelWorkbook -Path <file>` で作る。拡張子が `.xlsm` ならマクロ有効ブック、`.xlsx` なら通常のブックで保存する。既にファイルがあればエラーで止まる。作ったブックの絶対パスを出力するので、要らなければ `[void]` で捨てる。上限は `-TimeoutSeconds`（既定 120）。
- `Use-ExcelWorkbook` は `Invoke-ExcelScript.ps1` と同じ `-Save`、`-Visible`、`-EnableMacros`、`-EnableEvents`、`-TimeoutSeconds` を取り、上限と終了の扱いも同じ（SKILL.md の「COM 編集」）。
- モジュールを読み込むと、ScriptBlock の中で `Invoke-ExcelMacro` も使える（SKILL.md の「マクロを動かす」）。
- `Use-ExcelWorkbook` は ScriptBlock が出力した値をそのまま返す。COM メソッドの戻り値（`Worksheets.Add()` など）も混ざるので、要らないものは `[void]` で捨てる。COM オブジェクトを返すと、ブックを閉じた後も参照が残る。

## 別ブックからシートをコピーするとき

- コピー元を同じ Excel インスタンスで読み取り専用で開き、リンク更新を抑止する。`Worksheet.Copy` で作業コピーにシートを追加し、コピー元を開いたままコピー先を保存する。コピー元を先に閉じると保存時に Excel が終了し、`0x800706BE` が返った例がある。`Use-ExcelWorkbook -Save` は ScriptBlock の後で保存するため、ScriptBlock 内でコピー元を閉じる場合は、ScriptBlock 内でコピー先を保存してから、`finally` でコピー元を閉じる。
- 保存後にコピー先を開き直し、定義名の名前・参照先・スコープと、入力規則・書式を元シートと照合する。`Worksheet.Copy` が成功しても、コピー先のブックでフォントが変わる場合がある。
