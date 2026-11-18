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
- **コードを空にした文書モジュールは、書き出されない。** `Sheet1` は 0 行になり、`Export-Vba.ps1` が書き出したのは `documents/ThisWorkbook.cls` と `standard/Keep.bas` だけだった。照合は書き出しを比べるので、空にすれば、ブックにだけあるものから外れる。

正本からモジュールを消した時の手順は、`plugins/xlwright/skills/xlwright/references/vba.md` の「正本から消したモジュール」にある。その ScriptBlock を流した結果は、`vba-roundtrip.md` の「git diff --no-index で比べた時」にある。

## 未確認

- 取り込みの途中（いくつかを取り込んだ後）で失敗した時に、ブックが変わらないこと。保存しない作りからの見込み。ブックを開いた後、何かを取り込む前の失敗（ブックに無い文書モジュール）でブックが変わらないことは、試験の `input-rejected` が確かめる。
- 標準モジュール、クラス、UserForm、文書モジュールのほかの種別（ActiveX デザイナーなど）。
- マクロやイベントを有効にして開いたブックでの `Remove`。
- コントロールの多い UserForm、画像を持つ UserForm の `Remove`。
- VBA プロジェクトがロックされたブック。
- グラフシートの文書モジュール。
