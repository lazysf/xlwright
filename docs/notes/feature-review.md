# スキルが持つ機能と持たない機能の理由

- Date: 2026-11-19
- 出典: 人と決めた判断（2026-11-17、2026-11-18、2026-11-19）、commit `9c9d647`、`e93d79f`、`98358d0`

機能を置くかは、目的（Excel を土台にしたツールの開発、AGENTS.md）に当たるかと、ほかの機能を呼んで膨らまないかで決める。機能の一覧と使い方は、`plugins/xlwright/skills/xlwright/SKILL.md` と `references/` を正本とする。2026-11-18 に、その時のすべての機能が目的に当たることを確かめた。

## 置いている機能のうち、理由の要るもの

- `Test-VbaCompile.ps1`、`Invoke-ExcelMacro`、`Read-ExcelRange` と `Write-ExcelRange`、`Save-WorkbookScreenshot.ps1`、`Get-WorkbookSummary.ps1` は、代わりの手段が無いか、無いと誤りが出る。
- `Find-CellError.ps1` と `Test-Prerequisite.ps1` は、無くても困らないが、ほかの機能を複雑にしていない。
- `Get-WorkbookSummary.ps1` と `Find-CellError.ps1` は、受け取ったブックの分析にも使える形をしている。置いているのは、手を入れる前に既存のツールのブックを知るためと、シートの構成を変えた後にツールの数式が壊れていないかを見るためである。
- 試験の `-Changed` は、手で項目を選ぶと遅く、選び誤るという、実際の作業から出た理由がある。

## 置いていない機能

- **ブックに置いた VBA のテストをまとめて動かすスクリプトと、配布ブックからテストを外すスイッチ。** 確かめるマクロは、ScriptBlock の中で `Invoke-ExcelMacro` を並べれば動かせる。テストをブックに置くと、配布ブックから外す仕組みが要るようになる。ブックの `Test_` の `Sub` をすべて動かす作りでは、自分で書いていないマクロも動く。
- **正本とブックを照合するスクリプトと、正本に無いコンポーネントを一括で消すスイッチ。** VBE が整えた正本は往復しても変わらないので、照合は `Export-Vba.ps1` の書き出しと `git diff --no-index` で足りる（[vba-roundtrip.md](./vba-roundtrip.md)）。消すのは正本からモジュールを消した時だけで、名前を指定した `VBComponents.Remove` で足りる。一括で消すと、人がブックに足したコードまで消すおそれがある。それを防ぐには照合で対象を調べる手順が要り、機能が膨らむ。手順と注意は [references/vba.md](../../plugins/xlwright/skills/xlwright/references/vba.md) の「正本との照合」と「正本から消したモジュール」にある。

## 扱うブックの形式

扱うのは `.xlsx`、`.xlsm`、`.xlsb`、`.xlam` の 4 つで、それ以外はどのスクリプトもブックに触れる前に拒否する。人と決めた判断（2026-11-19）。形式ごとに使える場面は、`SKILL.md` の「扱うブックの形式」を正本とする。

- `.xlsm` はツールの本体で、作業ブックと配布の既定。
- `.xlsx` は、ツールが読み書きするブックやひな形と、シートだけのツールのために置く。
- `.xlam` は、アドインとして配るために置く。作業ブックは `.xlsm` か `.xlsb` のままにし、配布の時だけ `.xlam` にする（[publish-addin.md](./publish-addin.md)）。出力を読み戻して確かめるので、開く対象には入れる。
- `.xlsb` は、ツールの本体を `.xlsb` で作る利用者のために置き、`.xlsm` と同じ場面で使える。スキルのスクリプトを一通り `.xlsb` に通して確かめた。UserForm を持つ `.xlsb` は確かめていない（[xlsb.md](./xlsb.md)）。
- `.xltx` と `.xltm` は、`Workbooks.Open` で開くとテンプレートそのものが開く。その扱いを確かめていないので外した。
- `.xls` は customUI を持てないので外した。

拒否するのは、確かめていない形式で、スクリプトが中途半端に動くのを避けるためである。一覧は `XlWright.psm1` の 1 箇所にだけ置く。

## 確かめる画面の拡大率

画像を扱うスクリプトは、画面の拡大率 100% でだけ確かめる。人と決めた判断で、100% で確かめれば足りるとした。100% 以外の時、`Save-WorkbookScreenshot.ps1` は `Warnings` に警告を出す。`Test-RibbonLayout.ps1` の `-ScreenshotDir` は警告を出さず、タブの領域がウィンドウの画像からはみ出した時だけ、そのタブの `Error` に拡大率を疑う理由を入れる。確かめていない範囲は、`references/com.md` と `references/ribbon.md` に書いてある。

`description` を今の形にしている理由は、[skill-description.md](./skill-description.md) にある。
