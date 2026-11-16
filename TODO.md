# TODO

スキル xlwright の改善と追加機能の案を、優先度の順に書きます。
どれも未着手の案で、実際の Excel での実現性は確かめていません。スクリプト名とスイッチ名は仮のものです。

着手する時は、実験が要る項目は先に実験し、結果を `docs/notes/` に残します。終えた項目はここから消し、利用者に見える変更は [CHANGELOG.md](CHANGELOG.md) に書きます。

## 優先度 1: 書いた VBA が動くかを確かめる

今の VBA の検証は、コードの読み戻し、コンパイル（`Test-VbaCompile.ps1`）、`Invoke-ExcelMacro` やリボンのボタンで動かすこと、VBA で書いたテストの実行（`Invoke-VbaTest.ps1`）です。ここが一番弱いので、最初に手を付けます。

- [ ] 中断モードに入った VBE を戻せるかの実験。戻せれば、マクロの実行中のコンパイルエラーでも Excel を止めずに済む（[docs/notes/macro-dialog.md](docs/notes/macro-dialog.md)）。

## 優先度 2: 正本との同期と、毎回の手間

- [ ] 正本とブックの同期の確認（`Compare-Vba.ps1`）。ブックから書き出して正規化し（先頭の空行、識別子の大文字と小文字）、差のあるコンポーネントを JSON で返す。読み取りだけを行う。
- [ ] 取り込みの削除同期。`Import-Vba.ps1` に `-Prune` のような明示のスイッチを足し、正本に無いコンポーネントをブックから消す。既定の動きは変えない。
- [ ] 前提の確認を 1 本にまとめる（`Test-Prerequisite.ps1`）。Excel のプロセス、`AccessVBOM`、WMI、Git の状態を読み取りだけで調べ、JSON で返す。
- [ ] ブックの概要の取得。シートの一覧、`UsedRange`、定義名、テーブル、外部リンク、データ接続、VBA のコンポーネント、customUI の有無を 1 回で JSON にする。
- [ ] 2 次元配列を一括で読み書きする helper を `XlWright.psm1` に置く。`Cells.Item` への代入が型のキャストで落ちることがある件と、セルごとの COM 呼び出しを避ける。
- [ ] エラーセルの走査。編集の後に `#REF!` や `#NAME?` を `SpecialCells` で列挙する。

## 優先度 3: 見た目の確認と、扱える範囲

- [ ] シートの見た目の確認。範囲を PNG にして、書式、条件付き書式、グラフ、列幅を画像で見る。`PrintWindow` と `Range.CopyPicture` のどちらが使えるかを実験する。
- [ ] UserForm の実行時の画像。表示したフォームを撮り、レイアウトを確かめる。
- [ ] リボンの画像の埋め込み。`Add-CustomUI.ps1` で `image` 属性と `customUI/images` を扱う。
- [ ] リボンの押下の対象を広げる。toggleButton、menu の項目、dropDown。
- [ ] Backstage と contextMenu の扱い。
- [ ] 参照設定の正本化。`VBProject.References` を `vba/references.json` に書き出し、取り込みの時に照合する。
- [ ] `.xlam` の出力。`Publish-Workbook.ps1` の出力先として、アドインを選べるようにする。

## 優先度 4: 確かめていないことを埋める

機能を足さずに進められます。ほかの項目の合間に進めます。

- [ ] `references/ribbon.md` で確かめていないとした場面を確かめる。画面の拡大率が 100% 以外での画像、簡易表示のリボンの見分け、無効なボタンを指定した時のエラー。
- [ ] UI Automation で押していた時にだけ確かめた場面を、MSAA で押す今の形で `-Full` の試験に足す。英字のラベル、ダイアログを出さないボタン、MsgBox を 2 回続けて出すボタン、押せない要素を指定した場合、30 秒動くボタン。
- [ ] 32-bit の Excel と、Microsoft 365 以外の版での扱い。`AccessVBOM` のレジストリのパスは、`SKILL.md` と `tests/Test-XlWright.ps1` で `16.0` に固定している。確かめるか、確かめていないと明記する。
- [ ] eval の case を足す。配布ブックの出力、VBA の取り込み、既存のマクロを確認なしに動かさないこと。
- [ ] `plugins/xlwright/evals/results/` にたまった結果を整理する（Git には入っていない）。
