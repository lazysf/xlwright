# TODO

全体の段階は [ROADMAP.md](ROADMAP.md)、実施順序と期待結果は [Spike-0](docs/spike-0.md#実施順序と着手条件) を正本とします。
TODOへの記載は、未承認の設計変更や実機操作を承認するものではありません。

## 直近のTODO

R3を対象とします。チェックは、記載した完了条件を満たしたときに付けます。
進め方は[R3の進め方](docs/spike-0.md#r3の進め方)、移行の判定と未解決の仮説は[T14の判定](docs/notes/spike-0-results.md#t14-poc-1への移行判定2026-11-01)を参照します。
実機の項目は、着手時に目的・対象・停止条件を示して個別に承認を得ます。

- [ ] **T16 6操作の公開契約を確定する。** 提供範囲、JSON Schema、envelope、省略時の挙動、状態遷移、未対応要求の拒否を確定し、承認を得ます。製品ツールはこの後に実装します。
- [ ] **T17 検証済みの処理をCoreへ移す。** Spike0の`ExecutionGate`・`OwnershipRecord`・`RecordTransition`・`EndSafety`・`StaRunner`を移し、非COMの確認をxUnitへ移します。T12のうち記録更新・期限・応答の確定・次回開始の判定は、Coreの非COMテストで確認します。
- [ ] **T18 製品の実Excel経路を用意する。** 固定コピーにサーバーとテストを含め、接続を維持するxUnitのシナリオを`npm run check`から除外します。固定コピー以外を実行させない検査の置き場を決めます。表示要求を送る手動復旧スクリプトも用意します。
- [ ] **T19 1セルの縦断を確認する。** session_start → range_set → range_get → session_save → session_end を、人が実行します（実機）。
- [ ] **T13 Claude Codeから縦断とS13を確認する。** Git管理外のlocal scopeで固定コピーを登録し、MCP起動経路からのCOM生成も確認します（実機）。
- [ ] **T20 6操作へ広げる。** PoC-1合格条件1〜9とL9・L10を確認します。既存Excelとの同時利用、所有Excelの予期しない終了、計算鮮度の契約の製品経路での検証、終了待ちが期限内に収まることを含みます（実機）。
- [ ] **T12 L4〜L8のExcel依存条件を確認する。** Quit後の残存、期限後の遅延完了、COM保留中のEOF、保存・Close等の遅延復帰を確認します（実機）。
- [ ] **開発環境の後始末。** 古いSpike0.dllと`src/XlWright.McpServer/bin/`を整理します。`.xlwright/development/`の`prepared/t11-*`、`releases/`、`inspector/`にある古いSpike0.dllは検査を持たないため、実行しません。`.xlwright/mcp-checks/`の合成データ（約1.4GB）を消すかをユーザーが判断します。S14の作成領域に残った`exported-S11Module.bas`の扱いもユーザーが判断します。

## 更新方法

作業終了時にTODOを更新し、段階の進捗が変わった場合は [ROADMAP.md](ROADMAP.md) も更新します。
進めない項目には、理由と再開条件を付記します。
段階が完了したら、完了履歴はGitと結果記録に残し、直近のTODOを次の段階へ入れ替えます。
