# TODO

スキル xlwright で、これからやることを書きます。実装済みの機能の見直しと、新しい機能の案があります。
新しい機能の案は、どれも未着手で、実際の Excel での実現性は確かめていません。スクリプト名とスイッチ名は仮のものです。

どの案も、スキルの文書とスクリプトを読んで Claude が出したもので、利用者からの具体的な依頼から生まれたものではありません。各項目の「理由」は、その時の推測です。
着手する前に、理由がまだ成り立つか、「足りる形」で足りるかを人に確かめます。理由に当たらない作り込みは、「やらないこと」に入れておきます。

着手する時は、実験が要る項目は先に実験し、結果を `docs/notes/` に残します。終えた項目はここから消し、利用者に見える変更は [CHANGELOG.md](CHANGELOG.md) に書きます。

## 実装済みの機能の見直し

0.1.0 の後に足した機能のうち、ほかの機能を呼んで膨らんだものです。新しい機能より先に片付けます。どれも未リリース（CHANGELOG の `[Unreleased]`）なので、消しても利用者は困らず、今が一番安く済みます。

見直して残すと決めたもの（2026-11-17）: `Test-VbaCompile.ps1`、`Invoke-ExcelMacro`、`Read-ExcelRange` と `Write-ExcelRange`、`Save-WorkbookScreenshot.ps1`、`Get-WorkbookSummary.ps1` は、代わりの手段が無いか、無いと誤りが出る。`Find-CellError.ps1` と `Test-Prerequisite.ps1` は、無くても困らないが、ほかの機能を複雑にしていない。試験の `-Changed` は、手で項目を選ぶ遅さと誤りという、実際の作業から出た理由がある。

見直して消したもの（2026-11-18）: VBA のテストの実行（`Invoke-VbaTest.ps1`）と、配布ブックからテストを外す `Publish-Workbook.ps1 -ExcludeTests`。ScriptBlock の中で `Invoke-ExcelMacro` を並べれば足り、テストをブックに置くと配布ブックから外す仕組みが要るようになるため。既定ではブックの `Test_` の `Sub` をすべて実行し、自分で書いていないマクロも動かすことにもなっていた。
同じ日に、正本との照合（`Compare-Vba.ps1`）と、正本に無いコンポーネントを消す `Import-Vba.ps1 -Prune` も消した。VBE が整えた正本は往復しても変わらないので、照合は `Export-Vba.ps1` の書き出しと `git diff --no-index` で足りる（`docs/notes/vba-roundtrip.md`）。消すのは正本からモジュールを消した時だけで、名前を指定した `VBComponents.Remove` で足りる。一括で消すスイッチは、人が足したコードまで消すおそれと、それを防ぐために照合で対象を調べる手順を呼び込んでいた。手順と注意は `references/vba.md` の「正本との照合」と「正本から消したモジュール」にある。

## 今やる

理由が今の機能の穴に当たり、範囲が小さいものです。上の見直しの後に、上から順に進めます。

今は、この節に残っている項目はありません。

## 依頼が来るまで保留

使う依頼がまだ無いか、人の手が要るものです。依頼が来たら、理由を確かめ直してから着手します。

- [ ] リボンの画像の埋め込み。`Add-CustomUI.ps1` で `image` 属性と `customUI/images` を扱う。
  - 理由: 今のアイコンは組み込みの `imageMso` だけ（`references/ribbon.md`）。自前のアイコン（社のロゴ、`imageMso` に無い絵柄）を求められると応じられない。
  - 待つもの: `imageMso` で足りない依頼。
  - 足りる形: 人が渡した PNG を、ボタンの `image` 属性から参照できるように埋め込むこと。
  - やらないこと: アイコンを描くこと、`loadImage` や `getImage` の callback。
- [ ] リボンの押下の対象を広げる。toggleButton、menu の項目、dropDown。
  - 理由: `Invoke-RibbonButton.ps1` が探すのは、UI Automation で `Button` に見える要素だけ。toggleButton や menu の中の項目を押せるかは、確かめていない（`references/ribbon.md`）。toggleButton は `Test-RibbonLayout.ps1` で読めているので、`Button` に見えていれば今でも押せる。押せなければ、callback がつながっているかを押して確かめられない。マクロの本体は、callback から分けておけば `Invoke-ExcelMacro` で試せる（`references/ribbon.md`）ので、足りないのは「つながり」の確認だけ。
  - 待つもの: toggleButton や menu を持つリボンの依頼。来たら、まず今の `Invoke-RibbonButton.ps1` で押せるかを試す。
  - 足りる形: toggleButton と menu の中のボタンを押せること。dropDown は、使う依頼が来てから考える。
  - やらないこと: 押した後の状態（toggleButton の押下状態）を callback と照らすこと。
- [ ] `.xlam` の出力。`Publish-Workbook.ps1` の出力先として、アドインを選べるようにする。
  - 理由: 配布の形がブック（`.xlsm`）しかない。リボンとマクロを、どのブックでも使える形で配りたい依頼には応じられない。
  - 待つもの: アドインで配りたい依頼。
  - 足りる形: 今の配布の手順のまま、保存の形式だけを `.xlam` にできること。
  - やらないこと: アドインの導入（Excel の設定を変えるので、人が行う）。
- [ ] 画面の拡大率が 100% 以外での画像と、簡易表示のリボンの見分けを確かめる（`references/ribbon.md` の未確認）。
  - 理由: 画面の拡大率を 125% や 150% にしたノート PC は多く、そこで画像が欠けたり、簡易表示のリボンを見誤ったりしても、今は気づけない。
  - 待つもの: 人が試せる時。Windows の表示設定と Excel のリボンの表示を変える必要があり、エージェントは変えない。
  - 足りる形: 確かめた結果を文書に書くこと。直せないものは、確かめた範囲を書く。
