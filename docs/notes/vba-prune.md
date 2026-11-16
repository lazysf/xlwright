# ブックから VBA のコンポーネントを消す

- Date: 2026-11-17
- 出典: 一時フォルダの新しいブックで `VBComponents.Remove` と `CodeModule.DeleteLines` を呼んだ実測。Microsoft 365 Excel 64-bit 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818。

正本に無いコンポーネントをブックから消す時に、COM で何ができるかを確かめた。標準モジュール、クラス、UserForm は `VBComponents.Remove` で消えた。文書モジュールは消せず、コードを空にすることだけができた。

## 実験の手順

1. 新しい `.xlsm` に、標準モジュール 2 つ（`Keep`、`Old`）、クラス 1 つ（`OldClass`）、ラベル 1 つの UserForm（`frmOld`）を作る。`ThisWorkbook` と `Sheet1` にコードを書く。`Keep` は `Old` の関数を呼ぶ。保存して閉じる。
2. 非表示の Excel で開き直す。`ThisWorkbook` と `Sheet1` に `VBComponents.Remove` を呼ぶ。続けて `Old`、`OldClass`、`frmOld` に呼び、呼ぶたびに `VBComponents` を数える。`Sheet1` のコードを `CodeModule.DeleteLines` で空にする。保存して閉じる。
3. 開き直して、コンポーネントを数える。`Export-Vba.ps1` で書き出す。

マクロとイベントは無効のまま行った（`Use-ExcelWorkbook` の既定）。

## 結果

- **文書モジュールは `Remove` で消せない。** `ThisWorkbook` でも `Sheet1` でも、`ArgumentException`（`0x80070057`、「値が有効な範囲にありません。」）になった。コンポーネントもコードも変わらなかった。
- **標準モジュール、クラス、UserForm は、呼んだ時点で消える。** `Remove` の直後に `VBComponents.Count` が 1 つ減り、一覧からも消えた。確認のダイアログは出なかった。
- **消したモジュールを呼ぶコードが残っていても、`Remove` と保存は通る。** `Keep` は `Old.O()` を呼んだままで、エラーは出なかった。壊れた参照は、コンパイルするまで分からない。
- **保存して開き直しても、消したコンポーネントは戻らない。** 残ったのは `ThisWorkbook`、`Sheet1`、`Keep` だった。
- **コードを空にした文書モジュールは、書き出されない。** `Sheet1` は 0 行になり、`Export-Vba.ps1` が書き出したのは `documents/ThisWorkbook.cls` と `standard/Keep.bas` だけだった。`Compare-Vba.ps1` は書き出しを比べるので、空にすれば `workbook_only` から外れる。

## Import-Vba.ps1 の -Prune の形にした理由

使い方の正本は `plugins/xlwright/skills/xlwright/references/vba.md` の「書き出しと取り込み」にある。

- **明示のスイッチを付けた時だけ消す。** 一部のコンポーネントだけを別の作業先から取り込む使い方があり、既定で消すと残りがすべて消える。
- **文書モジュールは、コードを空にする。** 消せないが、空にすれば書き出されなくなる。消す対象が `Compare-Vba.ps1` の `workbook_only` と同じ集合になり、`-Prune` の後の照合が一致する。
- **消す前に一覧だけを出すスイッチは持たない。** `Compare-Vba.ps1` が同じ集合を、ブックを変えずに返す。
- **取り込みがすべて通った後に消す。** 途中で失敗すれば `Use-ExcelWorkbook` が保存せずに閉じるので、消した状態のブックは残らない。
- **コンパイルはしない。** 消したモジュールを呼ぶコードが残っても取り込みは通る。確かめるのは `Test-VbaCompile.ps1` の役目で、文書にその手順を書いた。

`tests/Test-XlWright.ps1` の `vba-prune` が、これを確かめる（2026-11-17）。

- `-Prune` を付けない取り込みでは、正本に無い標準モジュール 2 つ、UserForm、シートのコードが残った。
- `-Prune` を付けると、標準モジュールと UserForm が消え、シートのコードが空になり、照合が一致した。もう一度流すと、何も消さなかった。
- 取り込めない正本（ブックに無い文書モジュール）に `-Prune` を付けると失敗し、ブックのファイルは変わらなかった。この失敗は、何かを消す前に起きる。
- スイッチを見ずに消す写し、文書モジュールを空にしない写し、`Remove` を呼ばない写しは、どれも `vba-prune` で落ちた。

## 未確認

- 取り込みの途中（いくつかを取り込んだ後）や、消している途中で失敗した時に、ブックが変わらないこと。保存しない作りからの見込みで、試験には無い。
- 標準モジュール、クラス、UserForm、文書モジュールのほかの種別（ActiveX デザイナーなど）。
- マクロやイベントを有効にして開いたブックでの `Remove`。
- コントロールの多い UserForm、画像を持つ UserForm の `Remove`。
- VBA プロジェクトがロックされたブック。
- グラフシートの文書モジュール。
