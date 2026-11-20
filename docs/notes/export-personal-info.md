# 配布ブックの個人の情報を消す

- Date: 2026-11-20
- 出典: 一時フォルダの新しいブックで、`Workbook.RemovePersonalInformation` と `Workbook.RemoveDocumentInformation` を当てて保存し、COM と zip の中身を読んだ実測。Excel 16.0.20527.20196（Microsoft 365）、Windows 11、Windows PowerShell 5.1.26100.9818。値は記録せず、空か、合成の値か、それ以外かだけを見た。

`Export-Workbook.ps1` は、配布用の写しに `Workbook.RemovePersonalInformation = $true` を立ててから保存する。フラグはファイルに残るので、後のロックの保存でも消えたままになる。作業ブックには立てない。作業ブックの保存でも作成者などが消え、開発中のブックの扱いが変わるため。

`RemoveDocumentInformation(xlRDIDocumentProperties)`（`8`）は、タイトルまで消すので使わない。`RemoveDocumentInformation(xlRDIRemovePersonalInformation)`（`4`）は、`RemovePersonalInformation` を `true` にするのと同じ結果だった。`xlRDIAll`（`99`）は、ほかの内容まで消し得るので試していない。定数は Microsoft Learn の `XlRemoveDocInfoType` による。

## 確かめた事実

合成の値を入れた作業ブック（`.xlsm`、VBA の標準モジュール、メモ 1 件、スレッド形式のコメント 1 件）を Excel で保存すると、次の場所に個人の情報が残った。

| 場所 | 中身 |
| --- | --- |
| `docProps/core.xml` | `dc:creator`（作成者）、`cp:lastModifiedBy`（最終の編集者。保存した人の名前を Excel が書く） |
| `docProps/app.xml` | `Company`、`Manager` |
| `xl/workbook.xml` | `x15ac:absPath`（保存先のフォルダの絶対パス。home の下なら、ログインのユーザー名を含む） |
| `xl/comments1.xml` | メモの作成者（ログインのユーザー名） |
| `xl/persons/person.xml` | スレッド形式のコメントの作成者（`displayName`、`userId`、`providerId`） |

`xl/vbaProject.bin` には、ログインのユーザー名も home のパスも含まれていなかった。

`RemovePersonalInformation = $true` を立てて保存すると:

- `Author`、`Last Author`、`Manager`、`Company` が空になった。`Title` は残った。
- `x15ac:absPath` が消え、`xl/workbook.xml` に `filterPrivacy` が入った。読み戻した `RemovePersonalInformation` は `true` だった。
- メモの作成者は、ログインのユーザー名でない 3 文字に置き換わった（値は記録していない。Excel の表示の言語の「作成者」と見ているが、確かめていない）。`person.xml` の `person` は消え、COM で読むスレッド形式のコメントの作成者も 3 文字だった。メモとスレッド形式のコメントは 1 件ずつ残った（本文は読んでいない）。
- そのブックを開いて編集し、保存し直しても、上の項目は空のままだった。
- そのブックを `SaveAs` で `.xlam`（`55`）と `.xlsb`（`50`）にしても、空のままだった。`.xlsb` の `xl/comments1.bin` にも、ログインのユーザー名は無かった（UTF-8 と UTF-16 で探した）。処理しない `.xlsb` では、`person.xml` に作成者が残っていた。
- `.xlam` でマクロ（`AddNums`）が動いた。
- `Publish-Workbook.ps1`（今の `Export-Workbook.ps1`）で `.xlsm`、`.xlam`、`.xlsb` を出力すると、`Lock-VbaProject.ps1` の保存の後も、作成者、最終の編集者、管理者、会社は空で、`RemovePersonalInformation` は `true`、メモの作成者はログインのユーザー名でなかった。`.xlsm` の `xl/workbook.xml` に `absPath` は無かった。作業ブックの値は残った（`Test-XlWright.ps1 -Full -Case publish`。この試験のブックにスレッド形式のコメントは無い）。

## 未確認

- `-KeepPersonalInformation` の時に読む `Worksheet.CommentsThreaded` が、スレッド形式のコメントを持たない版の Excel にもあるか。Microsoft 365 の版でしか確かめていない。
- 画像、図形、ピボット、データ接続、プリンタの設定を持つブックでの挙動。今回のブックには無い。
- インクの注釈、`Workbook.Names` のコメント、シナリオのコメント。
