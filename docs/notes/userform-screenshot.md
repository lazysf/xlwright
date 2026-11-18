# 表示した UserForm を画像にする

- Date: 2026-11-18
- 出典: 一時フォルダの新しい `.xlsm` で、実験用スクリプトを流した実測。Excel 16.0.20527（Microsoft 365 64-bit、日本語の表示）、Windows 11、Windows PowerShell 5.1.26100.9818（STA）、画面の拡大率 100%（96 dpi）。

UserForm の見た目（文字が切れる、コントロールが重なる、フォントが効いていない）は、`Designer` のプロパティの数値からは分からない。表示したフォームを画像にできるかを確かめた。`Show` はモーダルで `$xl.Run` を止めるので、モードレスで表示した。

## 実験の手順

1. 試験の `uf-create`（`plugins/xlwright/skills/xlwright/tests/data/scripts/uf-create.ps1`）で、ラベル、フレームの中のチェック欄、ボタンを持つ `frmInfo`（240 × 160 pt）を作る。`UserForm_Initialize` は、ラベルに A1 の文字を入れ、チェック欄を入れる。
2. マクロを有効にして開き、A1 に `RuntimeA1` を書く。Excel のウィンドウを、`Save-WorkbookScreenshot.ps1` と同じく、どの画面にも入らない位置に置いてから表示し、1 秒待つ。
3. 一時の標準モジュールから、フォームをモードレスで表示する。1 回目は `frmInfo.Show vbModeless` だけ。2 回目は、`Load frmInfo` の後に `StartUpPosition = 0`、`Left` を画面の外、`Top = 0` にしてから `Show vbModeless`。
4. Excel のプロセスの、クラスが `ThunderDFrame` のウィンドウを探し、`PrintWindow`（`PW_RENDERFULLCONTENT`）で撮る。`Unload` で閉じる。ブックは保存しない。
5. 前面のウィンドウのプロセスを 50 ms ごとに記録した。

## 結果

| | 1 回目: `Show vbModeless` だけ | 2 回目: 位置を決めてから `Show vbModeless` |
| --- | --- | --- |
| `$xl.Run` が戻るまで | 35 ms | 41 ms |
| `StartUpPosition`、フォームのウィンドウの位置 | 1（親の中央）。画面の左端の (0, 291) に出た | 0。画面の外の (-2000, 0) |
| 画像 | 320 × 213 px。撮れた | 1 回目と同じ見た目で撮れた |
| 前面のプロセス | Claude と VS Code だけ | VS Code だけ |

- 画像には、`UserForm_Initialize` の後の状態が写った。ラベルは `RuntimeA1`、チェック欄は入っていた。日本語の表題と文字、フレーム、ボタンも写った。タイトルバーは、選ばれていない色で描かれた。
- 1 回目は、親の Excel が画面の外にあるのに、フォームは画面の中に出た。
- `Unload` の後、`ThunderDFrame` のウィンドウは無くなった。

## Save-WorkbookScreenshot.ps1 の -UserForm の形にした理由

形は、2026-11-18 に人と決めた（新しいスクリプトにせず、既存のスクリプトに足す）。使い方の正本は `plugins/xlwright/skills/xlwright/references/com.md` の「見た目を画像にする」にある。

- **既存のスクリプトに足す。** Excel を画面の外に置いて表示する手順、`PrintWindow`、見張り、結果の JSON を、そのまま使える。
- **`-EnableMacros` を求める。** フォームのコード（`UserForm_Initialize`）が動くので、マクロを動かすことを呼ぶ側に明示させる。
- **表示する前に、位置を画面の外にする。** 上の 1 回目のとおり、`StartUpPosition` が親の中央のままだと、画面に出た。
- **`VBA.UserForms.Add` にフォームの名前を渡す。** フォームの名前を一時モジュールのコードに埋め込まずに済む。名前は、ブックのコンポーネントの中の UserForm（種別 3）であることを、先に確かめる。
- **一時モジュールは保存しない。** ブックは保存せずに閉じる。

`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` の `screenshot`（`-Full`）が、これを確かめる（2026-11-18）。何を確かめるかは、試験のファイルを正本とする。

- スクリプトをわざと壊した次の写しは、どれも `screenshot` で落ちた。`-EnableMacros` を見ない写し、フォームのウィンドウでなく Excel のウィンドウを撮る写し、種別を確かめない写し。種別を確かめない写しは、`VBA.UserForms.Add` が失敗した後、呼び出しが約 2 分（見張りの上限の 120 秒）止まった。
- 位置を画面の外にしない写しは、この試験では見分けられない。フォームが画面に出るだけで、画像と結果は変わらない。

## 未確認

- 1 回目にフォームが画面の中に出た理由。親の中央が画面の外になるので、画面の中に寄せられたと見ている。
- フォームが、Excel のウィンドウを撮る `Save-WorkbookScreenshot.ps1`（`-UserForm` なし）の画像に写るか。別のウィンドウなので、写らない見込み。
- 種別を確かめない写しで、呼び出しを止めたもの。VBA の実行時エラーのダイアログが `$xl.Run` を止めたと見込んでいる。
- `UserForm_Initialize` や `UserForm_Activate` が MsgBox を出すフォームや、実行時エラーになるフォーム。モードレスでも `$xl.Run` が止まり、見張りの上限まで待つ見込み。
- 画面の拡大率が 100% 以外の時、複数の画面がある時。
- 画像やコントロールの多いフォーム、フォームが別のフォームを開く場合。
- `VBA.UserForms.Add` で作ったインスタンスと、既定のインスタンス（`frmInfo` と書いた時のもの）の違い。
