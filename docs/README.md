# 開発文書

ここにある文書を、xlwrightの現在の要件・設計・開発手順の正本とします。
設計上の採用と、実装・実機検証の完了は区別します。

- [requirements.md](requirements.md): 目的、対象環境、機能範囲、初期実用版の完了条件。
- [architecture.md](architecture.md): 構成、正本、実行制御、安全性、未確定の設計事項。
- [api.md](api.md): P0-coreの操作契約と共通規約。
- [development.md](development.md): セットアップ、配置方針、開発と検証の手順。
- [../ROADMAP.md](../ROADMAP.md): 開発段階の進捗と完了条件への参照。
- [../TODO.md](../TODO.md): 直近の作業、未解決事項、完了の根拠。
- [spike-0.md](spike-0.md): 最初の実機検証の仮説、実施計画、結果記録、PoCへの移行条件。

- [adr/0001-initial-lifecycle.md](adr/0001-initial-lifecycle.md): 初期ライフサイクルと所有記録の採用判断。
- [adr/0002-development-mcp-foundation.md](adr/0002-development-mcp-foundation.md): R1でのMCP接続基盤の前倒し、開発用機能と製品配布の境界。

調査記録は仕様の正本とは区別します。

- [notes/plan-review.md](notes/plan-review.md): 最小反復を先行する方針の根拠、実装手段の比較候補、計画の確認事項。
- [notes/lifecycle-proposal.md](notes/lifecycle-proposal.md): T1の比較案、所有記録の更新順、T2の採用判断への参照と後続の未承認事項。実機検証の結果ではありません。
- [notes/verification-environment.md](notes/verification-environment.md): T3の環境確認、PIAとMCP SDKの版・依存・承認範囲、取得経路と未検証事項。
- [notes/spike-0-preparation.md](notes/spike-0-preparation.md): T4の初回コンソール・基準ブックの準備、ビルドと非COM自己チェック、実機未確認事項。
- [notes/spike-0-results.md](notes/spike-0-results.md): ユーザー実行とMCP経由の正常系・次回開始、手動復旧、COM生成失敗の調査、接続基盤の検証・設定・更新、確認範囲と再開条件。
- [notes/mcp-reload-study.md](notes/mcp-reload-study.md): 標準再読込みの利用経路とKitsune MCP 0.21.0の静的検査、評価を停止した理由、未確認範囲。

[PROJECT_OVERVIEW_V22.md](../references/PROJECT_OVERVIEW_V22.md) は初期の参考資料です。
正式文書に未移行の記述を、自動的に現行仕様として採用しないでください。
仕様を変更するときは該当する正式文書を更新します。
