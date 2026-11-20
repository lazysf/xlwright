# docs

開発に加わる人に向けた、調べた事実と今の設計の理由を置きます。配布はしません。開発の手順は [DEVELOPMENT.md](../DEVELOPMENT.md) にあります。

- `docs/notes/`: 実測と調査の記録と、今の設計の理由。
  - [plugin-install.md](./notes/plugin-install.md): プラグインとして入れた時の振る舞いと、README の入れ方と版の決め方を今の形にしている理由。
  - [foreground.md](./notes/foreground.md): リボンの操作で Excel を前面に出さない方法と、その実測。
  - [ribbon-long-macro.md](./notes/ribbon-long-macro.md): リボンのボタンで長く動くマクロを、`Invoke-RibbonButton.ps1` が待つかの実測。
  - [vba-compile.md](./notes/vba-compile.md): VBA のコンパイルを COM から起こしてエラーを読む方法と、その実測。
  - [macro-dialog.md](./notes/macro-dialog.md): マクロの実行中に出たダイアログを見つけて閉じる方法と、その実測。
  - [macro-arguments.md](./notes/macro-arguments.md): `Invoke-ExcelMacro` が PSObject で包んだ引数の中の値を渡す理由と、その実測。
  - [vba-roundtrip.md](./notes/vba-roundtrip.md): 取り込んだ VBA を書き出した時に、正本と変わる所と、`git diff --no-index` で比べた時の出力の実測。
  - [vba-one-session.md](./notes/vba-one-session.md): VBA の取り込み、コンパイル、マクロの実行を 1 つの Excel で続けて行えるかの実測と、`Import-ExcelVba` などの関数に分けた理由。
  - [vba-prune.md](./notes/vba-prune.md): ブックから VBA のコンポーネントを消す時に COM でできることの実測。
  - [vba-references.md](./notes/vba-references.md): VBA の参照設定が書き出しと取り込みで引き継がれないことの実測。
  - [publish-addin.md](./notes/publish-addin.md): 配布の出力をアドイン（`.xlam`）にする方法と、作業ブックをブックのままにする理由と、その実測。
  - [publish-personal-info.md](./notes/publish-personal-info.md): 配布ブックに残る個人の情報の場所と、`RemovePersonalInformation` で消える範囲の実測。
  - [xlsb.md](./notes/xlsb.md): `.xlsb` をスキルのスクリプトに順に通した実測。
  - [prerequisite.md](./notes/prerequisite.md): 作業の前提（Excel のプロセス、`AccessVBOM`、WMI、Git）を読み取りだけで調べる方法と、その実測。
  - [workbook-summary.md](./notes/workbook-summary.md): シート、定義名、外部リンク、データ接続、VBA、customUI を COM と zip から読む方法と、その実測。
  - [range-value.md](./notes/range-value.md): セルの範囲を 2 次元配列で読み書きする時の型と形の実測と、`Read-ExcelRange` と `Write-ExcelRange` の形にした理由。
  - [cell-error.md](./notes/cell-error.md): エラー値のセルを `SpecialCells` と `Value2` で列挙する方法の比較と、`Find-CellError.ps1` の形にした理由。
  - [screenshot.md](./notes/screenshot.md): ブックの見た目を画像にする方法（ウィンドウの `PrintWindow` と `Range.CopyPicture`）の比較と、その実測。
  - [userform-screenshot.md](./notes/userform-screenshot.md): 表示した UserForm を、画面に出さずに画像にする方法の実測。
  - [excel-quit.md](./notes/excel-quit.md): `Quit` の後に Excel が終わるまでの時間の実測と、終わらなかった Excel が起動の時にしていたこと、`Stop-ExcelSession` が 5 秒待って止める理由。
  - [dev-cycle-timing.md](./notes/dev-cycle-timing.md): ツールのブックを作る 1 周（シート、VBA の取り込み、コンパイル、マクロ、読み戻し、エラー値のセル）で、スクリプトの呼び出しごとに時間がどこに掛かるかの実測。
  - [feature-review.md](./notes/feature-review.md): スキルが持つ機能と持たない機能を、目的に照らして今の形にしている理由。
  - [skill-description.md](./notes/skill-description.md): スキルの `description` を今の形にしている理由と、選ばれる範囲の実測。
  - [eval.md](./notes/eval.md): この repo で `claude plugin eval` がどう振る舞うかと、case を今の形にしている理由。
  - [codex-sandbox.md](./notes/codex-sandbox.md): Codex の sandbox の中で xlwright のスクリプトと COM の Excel が動くかと、Codex のエージェントで xlwright を使った時の振る舞いの実測。
  - [codex-exec.md](./notes/codex-exec.md): `codex exec` で eval の case を流す時の依頼文の足し方と記録の読み方の実測と、`Invoke-CodexCheck.ps1` の採点を今の形にしている理由。
  - [codex-plugin-install.md](./notes/codex-plugin-install.md): Codex で手元のフォルダからプラグインを入れた時の写しの置き場と、手直しが届く条件と、`.agents/skills` のリンクで手直し中のスキルを読ませた実測と、Claude 用と Codex 用の manifest を並べた時にどちらが使われるかの実測。
