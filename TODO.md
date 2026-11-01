# TODO

全体の段階は [ROADMAP.md](ROADMAP.md)、実施順序と期待結果は [Spike-0](docs/spike-0.md#実施順序と着手条件) を正本とします。
TODOへの記載は、未承認の設計変更や実機操作を承認するものではありません。

## 直近のTODO

R3を対象とします。チェックは、記載した完了条件を満たしたときに付けます。
進め方は[R3の進め方](docs/spike-0.md#r3の進め方)、移行の判定と未解決の仮説は[T14の判定](docs/notes/spike-0-results.md#t14-poc-1への移行判定2026-11-01)を参照します。
実機の項目は、着手時に目的・対象・停止条件を示して個別に承認を得ます。

### 停止時点の状態

R3のT18を終えた時点で、2026-11-01に作業を止めています。再開の入口は、下記の「T19の前提を実装する」です。

- コード: `095ef99`。この後の変更は文書だけです。
- 最後に通った`npm run check`: run `b2d3e9dcc49d4eb19946a0f828719577`、テスト131件。ソースは`095ef99`から変わっていません。
- `origin`へは一度もpushしていません。mainの全commitはこの端末にしかありません。pushは人がterminalで行います。
- 2026-11-01の照会で、EXCELとXlWright.McpServerのプロセスは0件でした。

### 項目

- [x] **T16 6操作の公開契約を確定する。** 提供範囲、JSON Schema、envelope、省略時の挙動、状態遷移、未対応要求の拒否を確定し、承認を得ます。製品ツールはこの後に実装します。2026-11-01にユーザーが承認し、[API契約](docs/api.md#poc-1の提供範囲)、[ADR 0005](docs/adr/0005-error-as-json-text.md)、[ADR 0006](docs/adr/0006-project-json-config.md)へ反映しました。SDKでの成立は[未確認](docs/api.md#未確認)のとおりT17以降で確認します。
- [x] **T17 検証済みの処理をCoreへ移す。** Spike0の`ExecutionGate`・`OwnershipRecord`・`RecordTransition`・`EndSafety`・`StaRunner`を移し、非COMの確認をxUnitへ移します。T12のうち記録更新・期限・応答の確定・次回開始の判定は、Coreの非COMテストで確認します。2026-11-01に`src/XlWright.Core/`へ移し、T16の契約のcode・detailに合わせました。StaRunnerは常駐の`StaThread`にしました。Spike0は変えていません。`npm run check`の全工程が通過しました（run `9fd1af6574e746a8a566f1f7f5b7ddc3`、テスト119件）。McpServerからCoreへの参照はまだありません。
- [x] **T18 製品の実Excel経路を用意する。** 固定コピーにサーバーとテストを含め、接続を維持するxUnitのシナリオを`npm run check`から除外します。固定コピー以外を実行させない検査の置き場を決めます。表示要求を送る手動復旧スクリプトも用意します。2026-11-01に、シナリオを別のテストプロジェクト`tests/XlWright.RealExcel.Tests/`に置き、`npm run prepare:product`と`scripts/Show-OwnedExcel.ps1`を加えました。`npm run check`の全工程が通過しました（run `2d05a04e53b84a409ee01743f9036843`、テスト131件）。エージェントが固定コピーを作り、コピーの接続シナリオが通ることを確認しました（Excelは起動しない）。表示要求のスクリプトは、合成の記録での拒否だけを確認しています。実機での実行は未実施です。
- [ ] **T19の前提を実装する。** 製品ツールは`dev_status`だけで、McpServerはCoreを参照していません。次をすべて満たしてからT19へ進みます。
  - 判断: Coreの公開範囲を決めます。Coreの型はすべて`internal`で、`InternalsVisibleTo`はテストだけを対象にしています。
  - 判断: 開発モード（`--development`、`dev_status`）を製品ビルドに含めるかを決めます。`Program.BuildHash`はsingle-file publishでは例外になります。
  - 実装: McpServerからCoreを参照し、T19が使う5ツールを実装します。
  - 確認: `npm run snapshot:tools`でツール定義のスナップショットを更新し、差分を確認します。引数をツール側で検査する形がSDK 2.3.0で組めるかを確かめ、[API契約の未確認](docs/api.md#未確認)を更新します。
  - ガード: `tests/XlWright.RealExcel.Tests/PreparedCopy.cs`の場所の確認を直します。リポジトリを`prepared`から導いて同じ`prepared`と比べているため、`.xlwright/development/prepared/product-*`の形ならどこに置いても通ります。
  - ガード: 製品シナリオに共有ガード`active-smoke.json`を作らせます。現在は存在しないことを確認するだけで、製品シナリオ同士やSpike0との同時実行を止められません。
  - ガード: `scripts/prepare-experiment.mjs`が出すmanifestを`PreparedCopy.Verify`に通すテストを加えます。現在`npm run check`はこのスクリプトを実行しません。
  - ガード: `scripts/build-guard.mjs`の鮮度の確認をCoreの変更にも広げます。現在はMcpServerのソースだけを見ています。
  - ガード: `scripts/Show-OwnedExcel.ps1`に`#requires -PSEdition Desktop`を加えます。PowerShell 7でも起動しますが、正しい記録まで拒否します。
- [ ] **T19 1セルの縦断を確認する。** session_start → range_set → range_get → session_save → session_end を、人が実行します（実機）。
- [ ] **T13 Claude Codeから縦断とS13を確認する。** Git管理外のlocal scopeで固定コピーを登録し、MCP起動経路からのCOM生成も確認します（実機）。
- [ ] **T20 6操作へ広げる。** PoC-1合格条件1〜9とL9・L10を確認します。既存Excelとの同時利用、所有Excelの予期しない終了、計算鮮度の契約の製品経路での検証、終了待ちが期限内に収まることを含みます（実機）。
- [ ] **T12 L4〜L8のExcel依存条件を確認する。** Quit後の残存、期限後の遅延完了、COM保留中のEOF、保存・Close等の遅延復帰を確認します（実機）。
- [ ] **開発環境の後始末。** 古いSpike0.dllと`src/XlWright.McpServer/bin/`を整理します。`.xlwright/development/`の`prepared/t11-*`、`releases/`、`inspector/`にある古いSpike0.dllは検査を持たないため、実行しません。`.xlwright/mcp-checks/`の合成データ（約1.5GB、`.xlwright/`全体で約2.1GB）を消すかをユーザーが判断します。S14の作成領域に残った`exported-S11Module.bas`の扱いもユーザーが判断します。

## 更新方法

作業終了時にTODOを更新し、段階の進捗が変わった場合は [ROADMAP.md](ROADMAP.md) も更新します。
進めない項目には、理由と再開条件を付記します。
段階が完了したら、完了履歴はGitと結果記録に残し、直近のTODOを次の段階へ入れ替えます。
