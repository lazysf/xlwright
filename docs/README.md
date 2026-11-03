# 開発文書

ここにある文書を、xlwrightの現在の要件・設計・開発手順の正本とします。
設計上の採用と、実装・実機検証の完了は区別します。

- [requirements.md](requirements.md): 背景、目的、目標、範囲、利用環境、利用の前提、安全性の要件。
- [architecture.md](architecture.md): 仕組みの方式、実行時の状態の置き場、部品の境界、未確認の仮説と成立しない時の扱い。
- [api.md](api.md): MCPのツールの共通の規約と一覧。各ツールの契約は段階ごとに加える。
- [development.md](development.md): セットアップ、配置方針、開発と検証の手順。
- [../ROADMAP.md](../ROADMAP.md): 開発の段階、到達する状態、完了の根拠。
- [../TODO.md](../TODO.md): 直近の作業、未解決事項、完了の根拠。

作り直しの前の判断（0001〜0009）の退役と、その時点で残した理由は、[ADR 0011](adr/0011-retire-pre-rebuild-adrs.md)にあります。その後の扱いは各行に書きます。

- [adr/0001-initial-lifecycle.md](adr/0001-initial-lifecycle.md): 初期ライフサイクルと所有記録の採用判断。0011で退役。
- [adr/0002-development-mcp-foundation.md](adr/0002-development-mcp-foundation.md): R1でのMCP接続基盤の前倒し、開発用機能と製品配布の境界。0011で退役。
- [adr/0003-inspector-centered-development.md](adr/0003-inspector-centered-development.md): Inspector中心の開発と、開発用のExcel開始ツールの削除。TODOの項目3で選び直す。
- [adr/0004-process-scoped-session.md](adr/0004-process-scoped-session.md): セッションをサーバープロセス単位に保ち、`session_id`を導入しない判断。新しい設計で採り直した。
- [adr/0005-error-as-json-text.md](adr/0005-error-as-json-text.md): 製品ツールの失敗を`McpException`の1行のJSONで返す判断。0007で改めた。
- [adr/0006-project-json-config.md](adr/0006-project-json-config.md): 利用プロジェクトの設定を`project.json`にする判断。0012で置き換え。
- [adr/0007-error-as-direct-result.md](adr/0007-error-as-direct-result.md): 製品ツールの失敗を`CallToolResult`で直接返す判断。新しいAPIで採り直した。
- [adr/0008-embedded-office-pia.md](adr/0008-embedded-office-pia.md): CoreのExcel型を既存PIAの埋込みで参照する判断。新しい設計で採り直した。
- [adr/0009-development-mode-without-excel.md](adr/0009-development-mode-without-excel.md): 開発モードを製品ビルドに含め、Excelを起動しない判断。新しい設計で、開発用のツールを持たない形に採り直した。

作り直しの後の判断です。

- [adr/0010-replace-all-modules.md](adr/0010-replace-all-modules.md): 標準・クラスモジュールを毎回すべて置き換える判断。
- [adr/0011-retire-pre-rebuild-adrs.md](adr/0011-retire-pre-rebuild-adrs.md): 作り直しで前提を失ったADRの退役と、退役させないADRの理由。
- [adr/0012-app-folder-convention.md](adr/0012-app-folder-convention.md): アプリのフォルダを規約で構成し、テストを写しにだけ取り込む判断。
- [adr/0013-read-and-edit-tools.md](adr/0013-read-and-edit-tools.md): ツールを読むツールと変えるツールに分ける判断。
- [adr/0014-synchronous-macro-run.md](adr/0014-synchronous-macro-run.md): マクロとテストの実行を、終わるまで応答を待つ形にする判断。
- [adr/0015-close-on-unknown-result.md](adr/0015-close-on-unknown-result.md): 結果不明では専用のExcelを終えてセッションを閉じる判断。

調査記録は仕様の正本とは区別します。

- [notes/verification-environment.md](notes/verification-environment.md): T3の環境確認、PIAとMCP SDKの版・依存・承認範囲、取得経路と未検証事項。
- [notes/spike-0-preparation.md](notes/spike-0-preparation.md): T4の初回コンソール・基準ブックの準備、ビルドと非COM自己チェック、実機未確認事項。
- [notes/spike-0-results.md](notes/spike-0-results.md): ユーザー実行とMCP経由の正常系・次回開始、手動復旧、COM生成失敗の調査、接続基盤の検証・設定・更新、確認範囲と再開条件、T7〜T15の実機結果、T14の移行判定、T16〜T18の準備の結果。
- [notes/mcp-reload-study.md](notes/mcp-reload-study.md): 標準再読込み・Kitsune・Hotloadの調査、Inspector中心の開発環境を採用する根拠、プラグインによる開発版の更新。
- [notes/mcp-server-development-handoff.md](notes/mcp-server-development-handoff.md): dotfilesへ持ち出すための、MCPサーバー開発環境の知見と標準化候補。単独で読める引継ぎ資料。
- [notes/mcp-best-practices.md](notes/mcp-best-practices.md): MCPサーバーの設計・開発の公式資料の調査、実測、現在の実装との差、Claude Codeでの呼出しと結果の制限。

[PROJECT_OVERVIEW_V22.md](../references/PROJECT_OVERVIEW_V22.md) は初期の参考資料です。
正式文書に未移行の記述を、自動的に現行仕様として採用しないでください。
仕様を変更するときは該当する正式文書を更新します。
