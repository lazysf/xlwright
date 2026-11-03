# 0005: 製品ツールの失敗を、McpExceptionの1行のJSONで返す

- Status: Accepted
- Amended-By: 0007

2026-11-01に、T16でユーザーが承認しました。
PoC-1の6操作の失敗は`isError: true`とし、`structuredContent`を付けません。
テキストは`error`と`session`を持つ1行のJSONで、`McpException`の文言として返します。形はAPI契約（`bdd2bb4`時点の`docs/api.md`。削除済み）にあります。

## 背景

API契約は、エラーに説明・セッション状態・`detail`・ログ参照先を含めると定めています。
[MCPのTools仕様](https://modelcontextprotocol.io/specification/draft/server/tools)は、`outputSchema`がある場合に`structuredContent`がそれに従うことを求めます。
同じ仕様のツール実行エラーの例は、`content`のテキストだけです。
C# SDK 2.3.0では、`McpException`の文言が応答に入り、他の例外の文言は入りません。[調査記録](../notes/mcp-best-practices.md#対象の版と実測)を参照してください。

## 検討した代替案

`CallToolResult`を直接返し、失敗も`structuredContent`に入れる案を比較しました。
この案では、`outputSchema`を成功と失敗の`oneOf`にする必要があります。
[実装基準](../development.md#mcpツールの実装基準)の`McpException`からも外れます。
採用した案は、成功の`outputSchema`だけで仕様を満たし、失敗の情報もJSONとして読めます。

## 影響

クライアントは、失敗の情報を`structuredContent`ではなくテキストから読みます。
`McpException`の文言がそのままテキストになることは、1箇所でだけ確認しました。全経路の保証ではありません。
対象のクライアントが失敗にも`structuredContent`を求めると分かったときは、この判断を見直します。
