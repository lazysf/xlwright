# MCPサーバーの設計・開発の調査

- Date: 2026-11-01
- 出典: [Tools仕様](https://modelcontextprotocol.io/specification/draft/server/tools)、[2026-07-28の変更点](https://modelcontextprotocol.io/specification/2026-07-28/changelog)、[Security Best Practices](https://modelcontextprotocol.io/specification/draft/basic/security_best_practices)、[Local Server Security](https://modelcontextprotocol.io/docs/2026-07-28/tutorials/security/local-server-security)、[Inspector CLI](https://modelcontextprotocol.io/docs/2026-07-28/tools/inspector/cli)、[Inspector Recipes](https://modelcontextprotocol.io/docs/2026-07-28/tools/inspector/recipes)、[mcpdo](https://modelcontextprotocol.io/docs/2026-07-28/tools/inspector/mcpdo)、[Protocol eras](https://modelcontextprotocol.io/docs/2026-07-28/tools/inspector/protocol-eras)、[Tasks拡張](https://modelcontextprotocol.io/extensions/tasks/overview)、[conformance](https://github.com/modelcontextprotocol/conformance)、[C# SDK](https://github.com/modelcontextprotocol/csharp-sdk)、[McpServerToolAttribute](https://csharp.sdk.modelcontextprotocol.io/api/ModelContextProtocol.Server.McpServerToolAttribute.html)、[Writing effective tools for agents](https://www.anthropic.com/engineering/writing-tools-for-agents)、ローカルの実行結果。

日々の手順と基準の正本は[開発ガイド](../development.md#mcpツールの実装基準)です。
本書は、その形にした理由と現在の実装との差を記録します。
調査では公式資料を優先し、第三者の解説記事は根拠に使っていません。

## 対象の版と実測

調査時点の最新仕様は2026-07-28です。これとは別にdraftがあります。
Toolsの要件はdraftの本文で確認しました。

2026-10-31に、`ModelContextProtocol.Core` 2.3.0のサーバーをInspector 2.13.0で確認しました。

- legacy・auto・modernの全世代で`dev_status`が成功しました。autoとmodernでは結果の`_meta`にサーバー情報が入りました。
- `isError`と未知のツールは、どちらも終了コード5でした。
- 削除前の`dev_excel_smoke_status`で、`ArgumentException("Unknown receipt.")`の応答は`An error occurred invoking 'dev_excel_smoke_status'.`でした。投げた文言は応答に含まれませんでした。
- 同じ箇所を一時的に`McpException`へ変えると、応答に`Unknown receipt.`が含まれました。
- 2026-11-01に、`McpServerTool`のサブクラスから投げた`McpException`を確認しました。応答のテキストは`An error occurred invoking '<tool>':`で始まりました。製品ツールは`CallToolResult`を直接返します。[ADR 0007](../adr/0007-error-as-direct-result.md)を参照してください。

## 開発フローへ取り入れた理由

自己修正ループには、Inspectorの[開発の進め方](https://modelcontextprotocol.io/docs/2026-07-28/tools/inspector/recipes#development-workflow)を反映しました。
公式の手順は、エラーの確認と、見つけた問題の自動検証への固定を含みます。

両プロトコル世代を検証するのは、2026-07-28で`initialize`の廃止などの大きな変更が入ったためです。
接続するクライアントの世代を、このサーバー側では選べません。

`tools/list`のスナップショットを持つのは、ツール定義がモデルへの指示になるためです。
公式資料は、定義の変更を注入の経路として扱っています。
公開インターフェースの変更を差分で確認できる効果もあります。

## 現在の実装との差

開発用の`dev_status`は、注釈を基準どおり明示しています。SDKの既定では`Destructive`と`OpenWorld`がtrueのためです。
結果はJSON文字列のテキストだけで、`outputSchema`はありません。開発用の識別情報なので、基準の対象とする製品ツールとは分けています。

[api.md](../api.md#共通規約)は、セッションをサーバープロセス単位で持ち、`session_id`を引数に取りません。プロトコルのセッションには依存しませんが、明示的なハンドルを渡す2026-07-28の案内とは異なります。
仕様は、プロトコルのセッションを廃止しました。
接続ごとの暗黙の状態に頼らず、明示的なハンドルを引数で渡す設計を案内しています。
この案内は非規範です。stdioのサーバーは1プロセスに1クライアントなので、現在の契約が直ちに動かなくなるとは確認していません。
契約は維持しています。理由は[ADR 0004](../adr/0004-process-scoped-session.md)にあります。

ハンドルへ変えても、stdioではサーバープロセスが終わると状態が失われます。
接続ごとに終了するInspector CLIでは、セッションをまたぐ操作を検証できません。

## 今は取り入れないもの

| 対象 | 理由 | 再検討の条件 |
| --- | --- | --- |
| conformance | 公式資料の実行方法がURL接続だけです。このサーバーはstdioだけです | stdio対応を確認したとき |
| mcpdo | 公式が実験的と明記しています。切断したstdioの子プロセスを次の利用時に自動で再接続するため、結果不明で再送しない方針との整合が未確認です。複数要求の検証は、接続を維持するxUnitのテストで行います | 接続を維持する手段をxUnitから替えるとき。再接続と再送の挙動を先に確認します |
| Tasks拡張 | 拡張であり、クライアントごとに対応が異なります。`execution_id`と状態照会で同じ形を実現できます | 対象クライアントの対応を確認したとき |
| エージェント評価ループ | 実際のAIクライアントかLLM APIの実行が必要です。現在は直接接続を登録していません | 製品ツールのツール選択を評価するとき |

## 未確認

- 対象のAIクライアントが使うプロトコル世代。
- mcpdoをxlwrightの開発サーバーへ接続した場合の挙動。実行していません。
