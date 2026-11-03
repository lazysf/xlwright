# 開発ガイド

本書は、xlwrightの開発の手順と確認の入口を記載します。
進め方の規則は [AGENTS.md](../AGENTS.md#開発の進め方) を正本とします。

## 進め方

- 作業は周期で進めます。周期の作るものと完了の条件は [TODO.md](../TODO.md) に書きます。
- Excelに触れない振る舞いは、テストを先に書いて失敗を確かめ、実装して通し、整えます。
- 実Excelの振る舞いは、捨てる前提の実験で確かめ、分かった振る舞いを結合テストで固定します。実験は`spikes/`に置きます。
- テストは、Excelを起動しないものと、実Excelを起動するものを分けます（[設計](architecture.md#本体の部品の境界)）。

## 確認の入口

今ある入口は次のとおりです。

すべての確認を1つのコマンドで行う:

```powershell
powershell -NoProfile -File scripts\check.ps1
```

[scripts/check.ps1](../scripts/check.ps1)は、次の順に進み、最初の失敗で止まります。

1. ビルドの識別子を決めます。形は[server_info](api.md#server_info段階0)の`build_id`です。
2. 依存を復元します。
3. サーバーを`.xlwright/builds/<build-id>/`へビルドします。
4. Excelを起動しないテスト（`tests/XlWright.McpServer.Tests`）を実行します。
5. 今回のビルドを`--no-excel`で起動し、Inspectorで`tools/list`と`server_info`を呼びます。`build_id`が今回の識別子と一致するかを確かめます。
6. Markdownをチェックし、`git diff --check HEAD`を行います。
7. 通ったビルドを`.xlwright/builds/latest.json`に記録します。

確認の入口を通らないビルド（`dotnet build`や`dotnet test`を直接実行した時）は、`build_id`が`unknown`になります。
テストはMicrosoft.Testing.Platformで実行します（[global.json](../global.json)）。直接実行する時は`dotnet test --solution xlwright.slnx`です。

Markdownのチェック:

```powershell
npm run lint:md
```

commitのhookと同じチェックを、全件に対して行う:

```powershell
prek run --all-files
```

hookの設定は [prek.toml](../prek.toml)、Markdownのルールは [.markdownlint-cli2.jsonc](../.markdownlint-cli2.jsonc) にあります。
