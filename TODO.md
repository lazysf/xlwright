# TODO

全体の段階は [ROADMAP.md](ROADMAP.md)、実施順序と期待結果は [Spike-0](docs/spike-0.md#実施順序と着手条件) を正本とします。
TODOへの記載は、未承認の設計変更や実機操作を承認するものではありません。

## 直近のTODO

R2を対象とします。チェックは、記載した完了条件を満たしたときに付けます。
T12（L4〜L8）とT13（S13）は[R3へ移しました](ROADMAP.md#後続工程で拾う事項)。

- [x] **T7 S10の起動設定を評価する。** 初期経路の設定を維持します。ActiveXは非対応のため対象外です。ForceDisableによるVBA抽出は`project_init`の実装前に残します。
- [x] **T8 S1〜S4のセル入出力を評価する。** 既存のセルAPI契約を維持します。[判定](docs/notes/spike-0-results.md#s4の限定整理とt8の判定)
- [x] **T9 S14の計算結果の鮮度を評価する。** 値取得前に明示計算・xlDone確認を行う設計を採用しました。[判定](docs/notes/spike-0-results.md#t9の完了判定と採用判断)
- [x] **T10 S9のSTAと再試行を評価する。** 限定評価です。実ExcelのRETRYLATER等は未検証です。[判定](docs/notes/spike-0-results.md#t10の限定完了判定2026-10-31)
- [x] **T11 S11の診断と終了拒否を評価する。** 限定評価です。約60秒の終了待ちと、人が閉じた後の残存の原因は未特定です。[判定](docs/notes/spike-0-results.md#t11-s11r2の全体の評価2026-11-01)
- [x] **T15 VBA経路を最小限確認する。** 遅延バインドのImportと、続くマクロ実行・保存・通常終了が成立しました。S5・S6・S7・S15はR4に残します。[結果](docs/notes/spike-0-results.md#t15-vba経路の最小確認2026-11-01)
- [ ] **T14 PoC-1への移行を判定する。** 未解決の仮説を整理し、[着手条件](docs/spike-0.md#poc-1へ進む前の確認)を照合して、R3のTODOを用意します。製品ツールを実Excelで検証する経路とセッション契約を最初に決めます。
- [ ] **開発環境の後始末。** 古いSpike0.dllと`src/XlWright.McpServer/bin/`を整理します。`.xlwright/development/`の`prepared/t11-*`、`releases/`、`inspector/`にある古いSpike0.dllは検査を持たないため、実行しません。`.xlwright/mcp-checks/`の合成データ（約1.4GB）を消すかをユーザーが判断します。S14の作成領域に残った`exported-S11Module.bas`の扱いもユーザーが判断します。

## 更新方法

作業終了時にTODOを更新し、段階の進捗が変わった場合は [ROADMAP.md](ROADMAP.md) も更新します。
進めない項目には、理由と再開条件を付記します。
段階が完了したら、完了履歴はGitと結果記録に残し、直近のTODOを次の段階へ入れ替えます。
