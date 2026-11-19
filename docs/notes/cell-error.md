# エラー値のセルを列挙する

- Date: 2026-11-17
- 出典: 一時フォルダの新しい `.xlsx` で、実験用スクリプトを非表示の Excel で流した実測。Microsoft 365 Excel 64-bit 16.0.20527.20196（日本語の表示）、Windows 11、Windows PowerShell 5.1.26100.9818。セルは `Write-ExcelRange` で書いた。

`#REF!` や `#NAME?` のセルを、`Range.SpecialCells` と、`Value2` を読んで調べる方法で列挙できるかを比べた。`SpecialCells` は、該当するセルが無い時と保護されたシートで同じ例外になり、エラーが多いと遅かった。`Value2` を読む方法は、どちらの問題も無かった。

## 実験の手順

1. シート `Data` に、数式のエラー（`=1/0`、`=NA()`、`=#REF!+1`、`=NOSUCHFN()`、ほかのセルのエラーを参照する `=A1`、無いシートを指す `=Sheet9!A1`）、定数のエラー（`#N/A` と `#DIV/0!` を打ち込んだ形）、エラーにならない数式（`=1+1`、`=IFERROR(1/0,0)`）、エラーをこぼす `=SEQUENCE(3)/0`、こぼれる先がふさがった `=SEQUENCE(2)` を置く。
2. エラーの無いシート、空のシート、エラーのある `veryHidden` のシート、エラーのある保護したシート、グラフシートを足す。
3. それぞれで `UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)` と `SpecialCells(xlCellTypeConstants, xlErrors)` を呼ぶ。
4. エラーの多いシートで、`SpecialCells` と、`UsedRange` を 10000 行ずつ `Value2` で読む方法の時間を比べる。

列挙の値は、Microsoft の `XlCellType`（`xlCellTypeFormulas` が -4123、`xlCellTypeConstants` が 2）と `XlSpecialCellsValue`（`xlErrors` が 16）の文書で確かめた（2026-11-17）。

## 結果

### SpecialCells

| 対象 | `SpecialCells(-4123, 16)`（数式） | `SpecialCells(2, 16)`（定数） |
| --- | --- | --- |
| `Data` | `A1,I1,K1,M1:N1,C3,E5,B10`（8 セル） | `G1:G2` |
| エラーの無いシート、空のシート | `COMException`「該当するセルが見つかりません。」 | 同じ |
| `veryHidden` のシート | `B2` | 例外（該当なし） |
| 保護したシート | `COMException`「保護されたシートに対して、このコマンドは使用できません。…」 | 同じ |
| グラフシート | `UsedRange` が無い（`null`） | |

- 該当するセルが無い時と、保護したシートの時は、どちらも HRESULT が `0x800A03EC` だった。文言は表示の言語で変わるので、例外からは見分けられない。
- Microsoft の `Range.SpecialCells` の文書は、該当するセルが無い時に何が起きるかを書いていない（2026-11-17 に確かめた）。
- こぼれた先のセル（`I2`、`I3`）は、値がエラーでも、数式にも定数にも入らなかった。`HasFormula` は `False`。こぼした元の `I1` だけが入る。
- 1 セル（`A2`）で呼ぶと、そのセルではなく、シートの使っている範囲全体から探した（結果は `Data` 全体と同じ）。2 セル以上の範囲（`A1:B2`、離れた範囲 `A1:A3,C3:C4`）では、その範囲の中だけから探した。
- 呼ぶ前後で、作業中のシートと選択範囲（`Data!A1`）は変わらなかった。

### Value2 で読む

- 保護したシートでも、`UsedRange.Value2` と `Formula2` は読めた。エラーのセルは、`Value2` で `Int32` のエラー値になる（[range-value.md](range-value.md)）。
- こぼれた先のセルも、`Value2` はエラー値になった。こぼした元は `SpillParent` で分かる（`I2` の `SpillParent` は `I1`、`I1` 自身も `I1`）。こぼれていないセルの `SpillParent` は `null`。
- 定数のエラーのセルは、`HasFormula` が `False` で、`Value2` がエラー値だった。

### 時間

| シート | 方法 | 時間 |
| --- | --- | --- |
| 20000 行 × 2 列、エラー 16667 セル（13333 の範囲） | `SpecialCells` | 599 ms |
| 同じ | `SpecialCells` の結果の範囲ごとに `Read-ExcelRange` | 40033 ms |
| 100000 行 × 10 列、エラー 142858 セル（7 セルに 1 つ、142858 の範囲） | `SpecialCells` | 61485 ms |
| 同じ | `UsedRange` を 10000 行ずつ `Value2` で読み、`Int32` を数える | 373 ms |

- `SpecialCells` の時間は、エラーのセルが離れて並び、範囲の数が多いほど延びた。
- 100000 行 × 10 列の数式を `Write-ExcelRange -Formula` で書くのに 31434 ms かかった。

## Find-CellError.ps1 の形にした理由

使い方の正本は `plugins/xlwright/skills/xlwright/references/com.md` の「エラー値のセルを探す」にある。

- **`SpecialCells` を使わず、`UsedRange` を `Value2` で読む。** `SpecialCells` は、該当なしと保護したシートを例外の HRESULT で見分けられず、こぼれた先のエラーを返さず、範囲が多いと遅かった。読んだ値は `Read-ExcelRange` で行の配列に直し、エラー値の名前も同じ表から取る。
- **10000 行ずつ、1 回で読むセルを 200000 までにして読む。** 100000 行 × 10 列を 10000 行ずつ読んで 373 ms だった。広いシートで 1 回の配列が大きくなりすぎないよう、セルの数でも区切る。区切りの大きさは測って決めたものではない。
- **`Formula` と `SpillParent` は、`Errors` に入れるセルだけで読む。** セル 1 つごとに COM を呼ぶので、`-MaxErrors` で回数を抑える。`SpillParent` は、こぼれていないセルでは失敗すると Microsoft の文書にあるので（2026-11-17 に確かめた）、先に `HasSpill` を見る。
- **計算はやり直さず、`Calculation` を返す。** やり直した時の時間と振る舞いは確かめていないので（「未確認」）、保存された値を読み、計算方法を添えて返す。計算方法の値は、Microsoft の `XlCalculation` の文書で確かめた（2026-11-17）。
- **エラーがあっても終了コードは 0。** 読めたかどうかと、エラーがあるかどうかを分ける。`Get-WorkbookSummary.ps1` と同じ。

`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` の `cell-error` が、これを確かめる（2026-11-17）。何を確かめるかは、試験のファイルを正本とする。

- スクリプトをわざと壊した次の写しは、どれも `cell-error` で落ちた。表示中のシートだけを読む写し、`HasSpill` を見ない写し、番地をシートの `A1` から数える写し、区切りの間の 1 行を飛ばす写し、`Truncated` を常に `false` にする写し、`-Sheet` の名前を確かめない写し、グラフシートも読む写し、`-MaxErrors` で打ち切らない写し。

## 未確認

- 数式を非表示にした（`FormulaHidden`）セルを含む、保護したシートでの `Formula2`。
- 手動の計算で保存されたブック。保存された値が古いと、エラーを見落とすか、直したエラーが残って見える。
- 外部のブックへのリンクを更新せずに開いた時の、リンクの数式の値。
- 計算をやり直した時の時間と振る舞い。
- エラーのセルの範囲の数が少ない時の、`SpecialCells` の時間。
- 16384 列に近い広いシート、1000000 行に近いシートでの時間とメモリ。
- 定義名、条件付き書式、入力規則の中の `#REF!`。セルの値には出ない。
- `.xlsm`、`.xlam`。`.xlsb` は、エラーの無いブックで `Count` が `0` になることだけを確かめた（[xlsb.md](./xlsb.md)）。ほかの形式（`.xls` など）は、スクリプトが開く前に拒否する。
