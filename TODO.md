# TODO

スキル xlwright の改善と追加機能の案を、優先度の順に書きます。
どれも未着手の案で、実際の Excel での実現性は確かめていません。スクリプト名とスイッチ名は仮のものです。

着手する時は、実験が要る項目は先に実験し、結果を `docs/notes/` に残します。終えた項目はここから消し、利用者に見える変更は [CHANGELOG.md](CHANGELOG.md) に書きます。

## 優先度 2: 見た目の確認と、扱える範囲

- [ ] シートの見た目の確認。範囲を PNG にして、書式、条件付き書式、グラフ、列幅を画像で見る。`PrintWindow` と `Range.CopyPicture` のどちらが使えるかを実験する。
- [ ] UserForm の実行時の画像。表示したフォームを撮り、レイアウトを確かめる。
- [ ] リボンの画像の埋め込み。`Add-CustomUI.ps1` で `image` 属性と `customUI/images` を扱う。
- [ ] リボンの押下の対象を広げる。toggleButton、menu の項目、dropDown。
- [ ] Backstage と contextMenu の扱い。
- [ ] 参照設定の正本化。`VBProject.References` を `vba/references.json` に書き出し、取り込みの時に照合する。
- [ ] `.xlam` の出力。`Publish-Workbook.ps1` の出力先として、アドインを選べるようにする。

## 優先度 3: 確かめていないことを埋める

機能を足さずに進められます。ほかの項目の合間に進めます。

- [ ] `references/ribbon.md` で確かめていないとした場面を確かめる。画面の拡大率が 100% 以外での画像、簡易表示のリボンの見分け、無効なボタンを指定した時のエラー。
- [ ] UI Automation で押していた時にだけ確かめた場面を、MSAA で押す今の形で `-Full` の試験に足す。英字のラベル、ダイアログを出さないボタン、MsgBox を 2 回続けて出すボタン、押せない要素を指定した場合、30 秒動くボタン。
- [ ] 32-bit の Excel と、Microsoft 365 以外の版での扱い。`AccessVBOM` のレジストリのパスは、`SKILL.md`、`scripts/XlWright.psm1`（`Get-ExcelPrerequisite`）、`tests/Test-XlWright.ps1` で `16.0` に固定している。確かめるか、確かめていないと明記する。
- [ ] eval の case を足す。配布ブックの出力、VBA の取り込み、既存のマクロを確認なしに動かさないこと。
- [ ] `plugins/xlwright/evals/results/` にたまった結果を整理する（Git には入っていない）。
