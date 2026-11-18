# 0.1.0 の後に足した機能の見直し

- Date: 2026-11-17、2026-11-18
- 出典: 人と決めた判断

0.1.0 の後に足した機能のうち、ほかの機能を呼んで膨らんだものを、新しい機能より先に見直した。どれも未リリース（CHANGELOG の `[Unreleased]`）だったので、消しても利用者は困らず、その時が一番安く済んだ。

## 残すと決めたもの（2026-11-17）

- `Test-VbaCompile.ps1`、`Invoke-ExcelMacro`、`Read-ExcelRange` と `Write-ExcelRange`、`Save-WorkbookScreenshot.ps1`、`Get-WorkbookSummary.ps1` は、代わりの手段が無いか、無いと誤りが出る。
- `Find-CellError.ps1` と `Test-Prerequisite.ps1` は、無くても困らないが、ほかの機能を複雑にしていない。
- 試験の `-Changed` は、手で項目を選ぶ遅さと誤りという、実際の作業から出た理由がある。

## 消したもの（2026-11-18）

- VBA のテストの実行（`Invoke-VbaTest.ps1`）と、配布ブックからテストを外す `Publish-Workbook.ps1 -ExcludeTests`。ScriptBlock の中で `Invoke-ExcelMacro` を並べれば足り、テストをブックに置くと配布ブックから外す仕組みが要るようになるため。既定ではブックの `Test_` の `Sub` をすべて実行し、自分で書いていないマクロも動かすことにもなっていた。
- 正本との照合（`Compare-Vba.ps1`）と、正本に無いコンポーネントを消す `Import-Vba.ps1 -Prune`。VBE が整えた正本は往復しても変わらないので、照合は `Export-Vba.ps1` の書き出しと `git diff --no-index` で足りる（[vba-roundtrip.md](./vba-roundtrip.md)）。消すのは正本からモジュールを消した時だけで、名前を指定した `VBComponents.Remove` で足りる。一括で消すスイッチは、人が足したコードまで消すおそれと、それを防ぐために照合で対象を調べる手順を呼び込んでいた。手順と注意は [references/vba.md](../../plugins/xlwright/skills/xlwright/references/vba.md) の「正本との照合」と「正本から消したモジュール」にある。

## 目的に照らした見直し（2026-11-18）

AGENTS.md に書いた目的（Excel を土台にしたツールの開発）に照らして、その時点のスキルの機能を見直した。0.1.0 の後に足して、まだ push していなかったものも含めた。目的から外れた機能は無く、消したものも無い。

| 機能 | ツールの開発での役目 |
| --- | --- |
| `Use-ExcelWorkbook`、`New-ExcelWorkbook`、`Invoke-ExcelScript.ps1` | ブックを作り、直す土台。Excel の起動と停止の安全も担う |
| `Read-ExcelRange`、`Write-ExcelRange` | ツールのシート（入力欄、設定、数式）を、型を崩さずに書く |
| `Export-Vba.ps1`、`Import-Vba.ps1`、`references.json`、`Test-VbaCompile.ps1`、`Invoke-ExcelMacro` | マクロの正本を管理し、自分が書いたマクロを確かめる |
| UserForm の手順、`Save-WorkbookScreenshot.ps1`（`-UserForm` を含む） | ツールの入力画面と、人にどう見えるかを確かめる |
| `Add-CustomUI.ps1`、`Test-RibbonLayout.ps1`、`Invoke-RibbonButton.ps1` | ツールの操作の入り口を作り、確かめる |
| `Lock-VbaProject.ps1`、`Publish-Workbook.ps1` | ツールを配る |
| `Test-Prerequisite.ps1`、`Get-WorkbookSummary.ps1` | 手を入れる前に、作業の前提と既存のツールのブックの中身を確かめる |
| `Find-CellError.ps1` | シートの構成を変えた後に、ツールの数式が壊れていないかを見る |
| 別ブックからシートをコピーする手順 | ひな形のシートを、ツールに取り込む |

外れていたのは、機能を説明する言い回しと eval の 1 件だった。どちらも、`description` の範囲を絞る TODO の項目で直す。

- `SKILL.md` の `description` と本文、`references/com.md` の言い回し（「シートの確認・編集」「複数のブックを順に扱う」など）が、受け取ったブックの分析や、多数のブックをまとめて処理する作業にも読める。
- eval の `sheet-check` が、受け取ったブックの集計の確認でスキルが選ばれることを期待している。
