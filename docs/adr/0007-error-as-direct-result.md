# 0007: 製品ツールの失敗をCallToolResultで直接返す

- Status: Accepted
- Amends: 0005

2026-11-01にユーザーが承認しました。
[ADR 0005](0005-error-as-json-text.md)の`McpException`の経路を、`isError: true`の`CallToolResult`を返す経路へ変えます。
応答の形は変えません。`structuredContent`を付けず、テキストは1行のJSONだけです。

## 背景

`McpServerTool`のサブクラスから投げた`McpException`を、2026-11-01にInspector 2.13.0で確認しました。
応答のテキストは`An error occurred invoking 'session_save':`で始まり、その後にJSONが続きました。
この接頭辞があると、テキスト全体をJSONとして読めません。

## 検討した代替案

`McpException`を維持し、クライアントに接頭辞を除かせる案を比較しました。
接頭辞の文面はSDKの実装に依存するため、API契約の一部にできません。

## 影響

`ProductTool`のサブクラスだけが直接返します。`dev_status`は従来どおりSDKに任せます。
