# ブックの概要を COM で読む

- Date: 2026-11-17
- 出典: 一時フォルダの実験用スクリプトを 1 回流した実測。Microsoft 365 Excel 64-bit 16.0.20527.20196（日本語の表示）、Windows 11、Windows PowerShell 5.1.26100.9818。

シートの一覧、使っている範囲、定義名、テーブル、外部リンク、データ接続、VBA のコンポーネント、customUI の有無を、非表示の Excel と zip から読めるかを確かめた。どれも読めた。customUI だけは、Excel がブックを開いている間は読めなかった。

## 実験の手順

1. 新しい `.xlsm` に、次を作って保存する。値とテーブル（`tblData`）と、別のブック（`other.xlsx`）を指す数式を持つシート `Data`。空のシート `Empty`。非表示のシート `Hidden`（`Visible` が 0）と `VeryHidden`（2）。グラフシート `Chart1`。CSV を `QueryTables.Add("TEXT;...")` で読み込んだシート `Query`。ブックの範囲の名前 `Rate` と、シートの範囲の名前 `Local`。標準モジュール 1 つ。
2. 閉じた状態で、`Add-CustomUI.ps1` でリボンを埋め込む。zip の項目を読む。
3. `Use-ExcelWorkbook`（マクロは無効、保存しない）で開き、zip をもう一度読む。`Sheets`、`Names`、`LinkSources`、`Connections`、`VBProject` を読む。
4. 何も入れていない `other.xlsx` でも、同じ項目を読む。

## 結果

### シート

| シート | `Type` | `Visible` | `UsedRange.Address` | `CountA(Cells)` |
| --- | --- | --- | --- | --- |
| `Data` | -4167 | -1 | `A1:E10` | 7 |
| `Empty` | -4167 | -1 | `A1` | 0 |
| `Hidden` | -4167 | 0 | `A1` | 0 |
| `VeryHidden` | -4167 | 2 | `A1` | 0 |
| `Query` | -4167 | -1 | `A1:B2` | 4 |
| `Chart1` | 3 | -1 | 例外 | 1 |

- 空のシートの `UsedRange` は `A1` になる。`A1` にだけ値があるシートと、番地では見分けられない。`WorksheetFunction.CountA(Cells)` が 0 かどうかで見分けられた。
- グラフシートは `UsedRange` が無く、番地を読むと例外になった。`CountA(Cells)` は 1 を返したので、グラフシートでは当てにならない。
- `Sheets` は、ブックの中の並びの順に返った。`CodeName` は、グラフシートでは `グラフ4` だった。
- テーブルは `ListObjects` から、名前と範囲（`tblData`、`A1:B3`）を読めた。

### 定義名

| `Name` | `RefersTo` | `Parent` |
| --- | --- | --- |
| `Rate` | `=Data!$B$2` | Workbook |
| `Data!Local` | `=Data!$A$2` | Worksheet |
| `Query!ExternalData_1` | `=Query!$A$1:$B$2` | Worksheet |

- `Workbook.Names` は、シートの範囲の名前も返す。その `Name` は `<シート名>!<名前>` の形になる。
- CSV の読み込みが、`ExternalData_1` を自動で作った。人が作っていない名前も混ざる。

### 外部リンク、データ接続

- `LinkSources(1)`（Excel のリンク）は、リンクが 1 つの時、要素 1 つの配列を返した。中身は `other.xlsx` のフルパスだった。リンクが無いブックでは `null` だった。
- `Connections` は、CSV の読み込みを、名前 `data`、`Type` が 4（テキスト）の接続として返した。接続の無いブックでは `Count` が 0 だった。
- リンク先のブックがある状態で開いた。`Use-ExcelWorkbook` はリンクを更新せずに開くので、ダイアログは出なかった。

### VBA

- `VBProject.VBComponents` から、名前、種別、行数を読めた。シートとグラフシートは、コードが無くても文書モジュール（種別 100）として並んだ。
- `Workbook.HasVBProject` は、標準モジュールのある `.xlsm` で `True`、何も入れていない `.xlsx` で `False` だった。`.xlsx` でも `VBComponents` は読め、空の文書モジュールだけが並んだ。

### customUI

- 閉じたブックの zip には、`customUI/customUI14.xml`、`xl/vbaProject.bin`、`xl/externalLinks/externalLink1.xml`、`xl/connections.xml` があった。
- Excel がブックを開いている間は、zip を読み取りで開けなかった（`IOException`、`0x80070020`）。zip は、Excel で開く前に読む。

## Get-WorkbookSummary.ps1 の形にした理由

使い方の正本は `plugins/xlwright/skills/xlwright/references/com.md` の「ブックの概要」にある。

- **シートの種別は、オブジェクトの型名（`Worksheet`、`Chart`）と `Type` の両方で決める。** グラフシートの `Type` は 3 だった。3 は、シートの種別（`XlSheetType`）では Excel 4.0 マクロシートの値で、グラフシートの値（-4109）ではない。`Type` だけでは見分けられない。値は Microsoft の `XlSheetType` の文書で確かめた（2026-11-17）。
- **空のシートの `UsedRange` は `null` にする。** Excel が返す `A1` のままだと、`A1` にだけ値があるシートと見分けられない。`CountA(Cells)` が 0 のシートを空とする。書式だけを付けたシートも空になる作りだが、確かめていない。
- **シートの範囲の名前は、`Scope` にシート名を入れ、`Name` から `<シート名>!` を外す。** 同じ名前がブックとシートの両方にある時に、どちらかを取り違えないため。
- **接続文字列と SQL は返さない。** サーバー名や認証情報が入っていることがあり、結果のファイルと会話に写る。種別の名前は、Microsoft の `XlConnectionType` の文書の値に合わせた（2026-11-17）。確かめたのは `text`（4）だけ。
- **VBA を読めない時も、ほかの項目は返す。** シートだけの作業や、VBA をロックした配布ブックの確認に使えるようにする。読めなかったことは `Vba.Components` の `null` と `Vba.Error` で分かる。空の配列（コードが無い）とは別にする。
- **customUI は、Excel で開く前に zip から読む。** 開いている間は zip を読めなかった。

`tests/Test-XlWright.ps1` の `workbook-summary` が、これを確かめる（2026-11-17）。実験に加えて、次を確かめた。

- `A1` にだけ値があるシートは `UsedRange` が `A1`、空のシートは `null` になった。
- 結果の JSON に、CSV のパス（接続文字列の中身）は入らなかった。
- 概要を取っても、ブックのファイルは変わらなかった。
- 何も入れていない `.xlsx` は、シート 1 つのほかは、どの項目も空だった。無いブックは、終了コード 1 で `Error` に理由が入った。
- 空のシートを見分けない写し、型名を見ない写し、名前からシート名を外さない写し、接続に項目を足した写し、コードの無い文書モジュールを省かない写し、`customUI/` を見ない写しは、どれも `workbook-summary` で落ちた。

VBA をロックしたブックの概要は、`-Full` の `publish` が確かめる（2026-11-17）。

- `Vba.Locked` が `true`、`Vba.Components` が `null` になり、`Vba.Error` に理由が入った。シートと customUI は返り、終了コードは 0 だった。
- ロックを見ずにコンポーネントを列挙する写しは、`publish` で落ちた。この写しでは、列挙が例外にならず、`Components` が配列で返った。ロックしたプロジェクトで列挙が何を返すかは、確かめていない。スクリプトは、列挙の前に `VBProject.Protection` を見る。

## 未確認

- リンク先のブックが無い時、リンクが複数の時。
- OLE DB や ODBC の接続、Power Query。`Type` の値と、接続文字列の読め方。
- 保護されたシートとブック、VBA プロジェクトがロックされたブック、`AccessVBOM` が 1 でない時。
- 非表示の定義名（`Visible` が `False`）。
- Excel 4.0 マクロシート、ダイアログシート。
- シートや定義名が数百あるブックでの時間。
- `.xlsx`、`.xlsm` のほかの形式（`.xlsb`、`.xls`）。
