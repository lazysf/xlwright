# UserForm に足したコントロールのフォント

- Date: 2026-11-20
- 出典: 一時フォルダの新しい `.xlsm` で、実験用スクリプトを `Invoke-ExcelScript.ps1` から流した実測。Excel 16.0.20527.20196（Microsoft 365、日本語の表示）、Windows 11、Windows PowerShell 5.1.26100.9818。

ユーザーからの報告では、このスキルでツールを開発中に既存のフォームにボタンとラベルを足してもらったら、同じダイアログの他の部分と違うフォントで足された。足したコントロールがどこからフォントを受けるかを、新しいブックで確かめた。

## 実験の手順

1. 新しいフォーム `frmA` を `VBComponents.Add(3)` で作り、PowerShell から `Designer.Controls.Add` でラベル `lblBefore` を足す。
2. 一時の VBA で、フォームの `Designer.Font` を `Meiryo UI`、11 にする。
3. PowerShell から `Controls.Add` でラベル `lblPs` とボタン `btnPs` を、一時の VBA から `Designer.Controls.Add` でラベル `lblVba` を足す。
4. 各段階で、一時の VBA からフォームと各コントロールの `Font` の `Name`、`Size`、`Bold`、`Charset` を `Designer` から読む。保存して閉じる。
5. マクロを有効にして開き直し、`Designer` と、`VBA.UserForms.Add` で読み込んだ実行時のフォームから、同じ値を読む。
6. 保存したブックの写しを人が Excel で開き、VBE のツールボックスからラベル `Label1` とコマンドボタン `CommandButton1` を手で置き、フォントには触らずに保存して閉じる。5 と同じ値を読む。

## 結果

| コントロール | 足した時 | 2 の後 | 開き直した後（デザイン時と実行時） |
| --- | --- | --- | --- |
| フォーム | MS UI Gothic, 9 | Meiryo UI, 11.25 | Meiryo UI, 11.25 |
| `lblBefore`（2 の前に足した） | MS UI Gothic, 9 | MS UI Gothic, 9 | MS UI Gothic, 9 |
| `lblPs`、`btnPs`（2 の後に PowerShell で足した） | Meiryo UI, 11.25 | | Meiryo UI, 11.25 |
| `lblVba`（2 の後に VBA で足した） | Meiryo UI, 11.25 | | Meiryo UI, 11.25 |
| `Label1`、`CommandButton1`（6 で VBE から手で足した） | | | Meiryo UI, 11.25 |

- `Bold` はすべて `False`、`Charset` はすべて 128 だった。
- 足したコントロールは、足した時のフォームの `Font` を写して受けた。PowerShell から足しても、VBA から足しても、VBE から手で足しても同じだった。
- フォームの `Font` を後から変えても、既にあるコントロールのフォントは変わらなかった。
- `Size` に 11 を入れると 11.25 になった。
- 一時の VBA は `finally` で消し、保存したブックに残らなかった。
- `Designer` の `Font` は、VBA からデザイン時に読めた。

## 分かったこと

コントロールのフォントがフォームの `Font` と違うフォームに足すと、足したコントロールだけがフォームの `Font` になり、他と違って見える。既存のコントロールのフォントを 1 つずつ変えた（`references/userform.md` の一時 VBA の例は、フォームでなくコントロールの `Font` を変える）フォームや、フォームの `Font` を変える前にコントロールを足したフォームで起きる。

`Controls.Add` の振る舞いが原因ではない。VBE で人が手で足しても、同じことが起きる。

## 未確認

- 報告のあったブックで、フォームの `Font` と既存のコントロールのフォントが違っていたか。そのブックは見ていない。
- `Italic` などの、読まなかったフォントの値も写して受けるか。
