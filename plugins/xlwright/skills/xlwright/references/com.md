# COM 編集の詳細

コマンド例の `<skill>` は、SKILL.md に書いたこのスキルのフォルダの絶対パスに置き換える。

## モジュールを直接使う

複数のブックを順に扱うなど、1 回の ScriptBlock に収まらない時は、スクリプトから `XlWright.psm1` を読み込む。スクリプトは一時領域に置き、`powershell.exe -NoProfile -File <script> -Path <workbook>` で実行する。

```powershell
param([string]$Path)
$ErrorActionPreference = 'Stop'
Import-Module '<skill>/scripts/XlWright.psm1'

Use-ExcelWorkbook -Path $Path -Save -ScriptBlock {
    param($xl, $wb)
    $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = 'Item'
    $ws.Range('B2').Value2 = [double]3
    $ws.Range('D2').Formula = '=B2*C2'
}
```

- **この形のスクリプトは ASCII だけで書く。** `powershell.exe`（5.1）は BOM なしの `.ps1` を UTF-8 として読まない。日本語のシート名、モジュール名、VBA コードは、UTF-8 のデータファイルに分けて `[IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)` で読み込む。この方法で、日本語のモジュール名、コメント、文字列リテラル（全角と半角カナを含む）が化けずに入り、`$xl.Run` の戻り値も一致することを確かめてある。日本語のパスやラベルを `-File` の引数として渡すのは問題ない（Git Bash から `-TabLabel 'ツール'` のように渡して動くことを確かめてある）。
- **日本語の結果はファイルに書いて読む。** 確認したい値は `[IO.File]::WriteAllLines($p, $lines, (New-Object Text.UTF8Encoding($false)))` で UTF-8 のファイルに書き出す。
- 新しいブックは `New-ExcelWorkbook -Path <file>` で作る。拡張子が `.xlsm` ならマクロ有効ブック、`.xlsx` なら通常のブックで保存する。既にファイルがあればエラーで止まる。作ったブックの絶対パスを出力するので、要らなければ `[void]` で捨てる。上限は `-TimeoutSeconds`（既定 120）。
- `Use-ExcelWorkbook` は `Invoke-ExcelScript.ps1` と同じ `-Save`、`-Visible`、`-EnableMacros`、`-EnableEvents`、`-TimeoutSeconds` を取り、上限と終了の扱いも同じ（SKILL.md の「COM 編集」）。
- `Use-ExcelWorkbook` は ScriptBlock が出力した値をそのまま返す。COM メソッドの戻り値（`Worksheets.Add()` など）も混ざるので、要らないものは `[void]` で捨てる。COM オブジェクトを返すと、ブックを閉じた後も参照が残る。

## 別ブックからシートをコピーするとき

- コピー元を同じ Excel インスタンスで読み取り専用で開き、リンク更新を抑止する。`Worksheet.Copy` で作業コピーにシートを追加し、コピー元を開いたままコピー先を保存する。コピー元を先に閉じると保存時に Excel が終了し、`0x800706BE` が返った例がある。`Use-ExcelWorkbook -Save` は ScriptBlock の後で保存するため、ScriptBlock 内でコピー元を閉じる場合は、ScriptBlock 内でコピー先を保存してから、`finally` でコピー元を閉じる。
- 保存後にコピー先を開き直し、定義名の名前・参照先・スコープと、入力規則・書式を元シートと照合する。`Worksheet.Copy` が成功しても、コピー先のブックでフォントが変わる場合がある。
