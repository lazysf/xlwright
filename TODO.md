# TODO

全体の段階は [ROADMAP.md](ROADMAP.md)、実施順序と期待結果は [Spike-0](docs/spike-0.md#実施順序と着手条件) を正本とします。
TODOへの記載は、未承認の設計変更や実機操作を承認するものではありません。

## 直近のTODO

R3を対象とします。チェックは、記載した完了条件を満たしたときに付けます。
進め方は[R3の進め方](docs/spike-0.md#r3の進め方)、移行の判定と未解決の仮説は[T14の判定](docs/notes/spike-0-results.md#t14-poc-1への移行判定2026-11-01)を参照します。
実機の項目は、着手時に目的・対象・停止条件を示して個別に承認を得ます。

### 停止時点の状態

T19を終えた時点で、2026-11-01に作業を止めています。再開の入口は、下記のT13です。

- コード: `2795c12`。この後の変更は文書だけです。
- 最後に通った`npm run check`: run `54258892720a4230bd6a73ec69a47005`、テスト271件。
- 固定コピー: `.xlwright/development/prepared/product-8cc5eca6afe1411c8dc55b0d868f688c/`。独立レビューの修正を含むT19の再実行に使いました。
- `origin`へは一度もpushしていません。mainの全commitはこの端末にしかありません。pushは人がterminalで行います。
- 同じ時点の照会で、EXCELとXlWright.McpServerのプロセスは0件でした。

### 項目

- [x] **T16 6操作の公開契約を確定する。** 提供範囲、JSON Schema、envelope、省略時の挙動、状態遷移、未対応要求の拒否を確定し、承認を得ます。製品ツールはこの後に実装します。2026-11-01にユーザーが承認し、[API契約](docs/api.md#poc-1の提供範囲)、[ADR 0005](docs/adr/0005-error-as-json-text.md)、[ADR 0006](docs/adr/0006-project-json-config.md)へ反映しました。SDKでの成立は[未確認](docs/api.md#未確認)のとおりT17以降で確認します。
- [x] **T17 検証済みの処理をCoreへ移す。** Spike0の`ExecutionGate`・`OwnershipRecord`・`RecordTransition`・`EndSafety`・`StaRunner`を移し、非COMの確認をxUnitへ移します。T12のうち記録更新・期限・応答の確定・次回開始の判定は、Coreの非COMテストで確認します。2026-11-01に`src/XlWright.Core/`へ移し、T16の契約のcode・detailに合わせました。StaRunnerは常駐の`StaThread`にしました。Spike0は変えていません。`npm run check`の全工程が通過しました（run `9fd1af6574e746a8a566f1f7f5b7ddc3`、テスト119件）。McpServerからCoreへの参照はまだありません。
- [x] **T18 製品の実Excel経路を用意する。** 固定コピーにサーバーとテストを含め、接続を維持するxUnitのシナリオを`npm run check`から除外します。固定コピー以外を実行させない検査の置き場を決めます。表示要求を送る手動復旧スクリプトも用意します。2026-11-01に、シナリオを別のテストプロジェクト`tests/XlWright.RealExcel.Tests/`に置き、`npm run prepare:product`と`scripts/Show-OwnedExcel.ps1`を加えました。`npm run check`の全工程が通過しました（run `2d05a04e53b84a409ee01743f9036843`、テスト131件）。エージェントが固定コピーを作り、コピーの接続シナリオが通ることを確認しました（Excelは起動しない）。表示要求のスクリプトは、合成の記録での拒否だけを確認しています。実機での実行は未実施です。
- [x] **T19の前提を実装する。** 2026-11-01に、公開範囲を`ExcelSession`の窓口に限り、開発モードを[ADR 0009](docs/adr/0009-development-mode-without-excel.md)、Excelの型参照を[ADR 0008](docs/adr/0008-embedded-office-pia.md)に決めました。5ツールとガード5件を実装し、スナップショットの差分とInspectorの拒否応答を確認しました。失敗の返し方は[ADR 0007](docs/adr/0007-error-as-direct-result.md)のとおりです。
- [x] **T19 1セルの縦断を確認する。** 2026-11-01に、人が固定コピーから実行し、数値1セルの読み書き・保存・終了が成功しました。[結果](docs/notes/spike-0-results.md#t191セルの縦断2026-11-01)を参照します。
- [ ] **T13 Claude Codeから縦断とS13を確認する。** Git管理外のlocal scopeで固定コピーを登録し、MCP起動経路からのCOM生成も確認します（実機）。
- [ ] **T20 6操作へ広げる。** PoC-1合格条件1〜9とL9・L10を確認します。既存Excelとの同時利用、所有Excelの予期しない終了、計算鮮度の契約の製品経路での検証、終了待ちが期限内に収まることを含みます（実機）。T19で送った`session_status`と`last_error`、`SERVERCALL_RETRYLATER`の再試行、タイムアウト時のウィンドウ診断もここで実装します。
- [ ] **T12 L4〜L8のExcel依存条件を確認する。** Quit後の残存、期限後の遅延完了、COM保留中のEOF、保存・Close等の遅延復帰を確認します（実機）。
- [ ] **開発環境の後始末。** 2026-11-01にユーザーの判断で、文書から参照されない生成物を削除しました（`.xlwright/`は約2.1GBから約350MB）。残りは、`.xlwright/spike0/fixture-preparation/`の`s11r2/`と`s14-*/`にある`exported-S11Module.bas`、用途が未確認の`.xlwright/fixture-builder/`と`.xlwright/t11-discarded/`です。これらの扱いはユーザーが判断します。

## 更新方法

作業終了時にTODOを更新し、段階の進捗が変わった場合は [ROADMAP.md](ROADMAP.md) も更新します。
進めない項目には、理由と再開条件を付記します。
段階が完了したら、完了履歴はGitと結果記録に残し、直近のTODOを次の段階へ入れ替えます。
