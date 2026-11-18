# TODO

スキル xlwright で、これからやることを書きます。項目は、[AGENTS.md](AGENTS.md) にある目的（Excel を土台にしたツールの開発）に当たるかで選びます。

着手する前に、理由がまだ成り立つか、「足りる形」で足りるかを人に確かめます。理由に当たらない作り込みは、「やらないこと」に入れておきます。

着手する時は、実験が要る項目は先に実験し、結果を `docs/notes/` に残します。終えた項目はここから消し、利用者に見える変更は [CHANGELOG.md](CHANGELOG.md) に書きます。「やらないこと」のうち、後で根拠として参照するものは、項目を消す前に `docs/notes/` か `references/` に書きます。項目と一緒に消えると、参照が切れるためです。

## 今やる

理由が、今の機能の穴か、スキルが選ばれる範囲の誤りに当たるものです。上から順に進めます。

- [ ] `SKILL.md` の `description` を、ツールの開発に絞る。`plugin.json` と `marketplace.json` の `description` もそろえる。
  - 理由: 今の `description` は「シートの値・数式・書式」を先に挙げていて、`xlsx` スキルなど通常の表計算のスキルと範囲が重なる。受け取ったブックの確認の依頼でも、スキルが選ばれることがある（2026-11-18 に試した `existing-macro` の case で、選ばれ方が揺れた）。
  - 足りる形: `description` に、目的（ツールの開発）と、使わない場面（受け取ったブックの分析、集計の確認、データの整形）を書く。`sheet-check` の依頼文を、ツールを作る途中の確認に替える。範囲外の Excel の依頼でスキルが選ばれないことを見る case を足し、eval を全件流す。着手する前に、使用量の見積もりを出す。
  - 公式 docs で確かめた制約（2026-11-18）: Claude Code はスキルの一覧で `description` と `when_to_use` を合わせて 1,536 文字で切るので、主な用途を先に書く（[skills](https://code.claude.com/docs/en/skills)）。`<` と `>` は `&lt;` と `&gt;` で表示される。Agent Skills の決まりでは、`description` は 1,024 文字まで、XML のタグを含めず、三人称で、何をするかといつ使うかを書く（[best practices](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/best-practices)）。今の `description` は 311 文字。
  - やらないこと: `xlsx` スキルの側を直すこと。範囲外の依頼を断る手順を `SKILL.md` に書くこと。

## 依頼が来るまで保留

使う依頼がまだ無いか、人の手が要るものです。依頼が来たら、理由を確かめ直してから着手します。

どの項目も、スキルの文書とスクリプトを読んで Claude が出した案で、利用者からの具体的な依頼から生まれたものではありません。各項目の「理由」はその時の推測で、実際の Excel での実現性は確かめていません。スクリプト名とスイッチ名は仮のものです。

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
