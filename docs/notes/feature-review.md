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
