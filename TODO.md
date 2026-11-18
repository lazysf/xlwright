# TODO

スキル xlwright で、これからやることを書きます。実装済みの機能の見直しと、新しい機能の案があります。
新しい機能の案は、どれも未着手で、実際の Excel での実現性は確かめていません。スクリプト名とスイッチ名は仮のものです。

どの案も、スキルの文書とスクリプトを読んで Claude が出したもので、利用者からの具体的な依頼から生まれたものではありません。各項目の「理由」は、その時の推測です。
着手する前に、理由がまだ成り立つか、「足りる形」で足りるかを人に確かめます。理由に当たらない作り込みは、「やらないこと」に入れておきます。

着手する時は、実験が要る項目は先に実験し、結果を `docs/notes/` に残します。終えた項目はここから消し、利用者に見える変更は [CHANGELOG.md](CHANGELOG.md) に書きます。

## 実装済みの機能の見直し

0.1.0 の後に足した機能のうち、ほかの機能を呼んで膨らんだものです。新しい機能より先に片付けます。どれも未リリース（CHANGELOG の `[Unreleased]`）なので、消しても利用者は困らず、今が一番安く済みます。

見直して残すと決めたもの（2026-11-17）: `Test-VbaCompile.ps1`、`Invoke-ExcelMacro`、`Read-ExcelRange` と `Write-ExcelRange`、`Save-WorkbookScreenshot.ps1`、`Get-WorkbookSummary.ps1` は、代わりの手段が無いか、無いと誤りが出る。`Find-CellError.ps1` と `Test-Prerequisite.ps1` は、無くても困らないが、ほかの機能を複雑にしていない。試験の `-Changed` は、手で項目を選ぶ遅さと誤りという、実際の作業から出た理由がある。

- [ ] 正本との照合のつながりを見直す。`Compare-Vba.ps1` → `Import-Vba.ps1 -Prune` → 「`-Prune` の前に `Compare-Vba.ps1` で消える対象を調べる」という手順、と膨らんだ。
  - 理由: VBE が整えた形の正本は、取り込んで書き出しても 1 バイトも変わらない（`docs/notes/vba-roundtrip.md`）。`Export-Vba.ps1` で一時フォルダに書き出し、`git diff --no-index` で比べれば、ほぼ足りる。足りないのは、UserForm の `.frm` の空行と、大文字小文字だけの差の見分け。`-Prune` が要るのは、正本からモジュールを消した時だけで、`VBComponents.Remove` の 1 行で代わりになる（`docs/notes/vba-prune.md`）。
  - 量: スクリプトとスイッチが約 190 行（`Compare-Vba.ps1`、`Import-Vba.ps1` の `-Prune`）、試験の項目が 2 つで約 95 行（`vba-compare`、`vba-prune`）。
  - 決めること: 次のどれにするかを人に確かめる。
    1. 両方消し、書き出しと `git diff --no-index` で比べる手順を `references/vba.md` に書く。`.frm` の空行と大文字小文字の差は、手順の注意として書く。
    2. `Compare-Vba.ps1` だけ残し、`-Prune` は消す。
    3. 今のまま残す。
  - 消す時に直すもの: `SKILL.md`（`-Prune` の注意）、`references/vba.md` の「書き出しと取り込み」と「正本との照合」、`docs/notes/vba-roundtrip.md` と `docs/notes/vba-prune.md`（実測は残し、スクリプトの形にした理由を消す）、README の「できること」、試験の `$CasesByFile`、CHANGELOG の `[Unreleased]`。

## 今やる

理由が今の機能の穴に当たり、範囲が小さいものです。上の見直しの後に、上から順に進めます。

- [ ] 参照設定の正本化。`VBProject.References` を `vba/references.json` に書き出し、取り込みの時に照合する。
  - 理由: `Export-Vba.ps1` と `Import-Vba.ps1` は、参照設定（Scripting Runtime など）を扱わない。正本の `vba/` から新しいブックを作ると、参照設定が無く、`Dictionary` などの型でコンパイルが通らない見込み（確かめていない）。正本だけでは、ブックを作り直せない。
  - 最初に: 参照設定を使うコードを新しいブックに取り込み、`Test-VbaCompile.ps1` が失敗することを確かめる。失敗しなければ、この項目は要らない。
  - 足りる形: 参照設定を書き出すこと。取り込む時に、正本にあってブックに無い参照設定を足すか、少なくとも警告すること。
  - やらないこと: 参照先のファイルの配布、版の違う参照設定の差し替え。
- [ ] UI Automation で押していた時にだけ確かめた場面を、MSAA で押す今の形で `-Full` の試験に足す。英字のラベル、ダイアログを出さないボタン、MsgBox を 2 回続けて出すボタン、押せない要素を指定した場合、30 秒動くボタン。あわせて、無効なボタンを指定した時のエラーも足す。
  - 理由: 押し方を MSAA に変えた後、これらの場面は試験に無い。今の形で壊れていても気づけない。無効なボタンのエラーは、`references/ribbon.md` に書いた振る舞いが実機で起きるかを確かめていない。
  - 足りる形: 既にある `ribbon` の項目に場面を足し、壊した写しで落ちること。`references/ribbon.md` の、確かめた場面を 2 つに分けた書き方を 1 つにまとめる。
  - やらないこと: 場面ごとに新しい項目を作ること。項目を増やすと、Excel を起動する回数と流す時間が増える。
- [ ] UserForm の実行時の画像。表示したフォームを撮り、レイアウトを確かめる。
  - 理由: 今の確認は、`Designer` のプロパティの数値と、`Load` した後の値だけ（`references/userform.md`）。ラベルの文字が切れる、コントロールが重なる、フォントが効いていない、といった見た目の崩れは、数値からは分からない。フォントは PowerShell から読めない。シートの画像（`Save-WorkbookScreenshot.ps1`）で UserForm が写るかは、確かめていない（`docs/notes/screenshot.md` の「未確認」）。
  - 最初に: 表示したフォームはモーダルで `Run` を止めるので、撮れるかを実験する。
  - 足りる形: `UserForm_Initialize` の後のフォームを 1 枚の PNG にして、Read で見られること。
  - やらないこと: 画像どうしの自動比較、フォームへの入力の操作。
- [ ] eval の case を足す。配布ブックの出力、VBA の取り込み、既存のマクロを確認なしに動かさないこと。
  - 理由: 今の eval は、スキルが選ばれるか、reference を読み分けるか、Excel を名前で止めないかだけを見る。既存のマクロを確認なしに動かさない規則は、`SKILL.md` の安全の規則なのに、eval が無い。`SKILL.md` を直した時に抜けても気づけない。
  - 最初に: 流す回数と使用量の見積もりを出して、人に確かめる。上の見直しで `SKILL.md` が変わるので、その後に作る。
  - 足りる形: Windows の eval はシェルを使えない（`DEVELOPMENT.md` の「eval」）ので、手順を立てる段階の振る舞いを見る case にする。

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
