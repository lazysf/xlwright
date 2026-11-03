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
7. 通ったビルドを`.xlwright/builds/latest.json`に記録し、開発版のプラグインの`plugins/xlwright-dev/.mcp.json`（Git管理外）を、そのビルドを起動する形で書き出します。

確認の入口を通らないビルド（`dotnet build`や`dotnet test`を直接実行した時）は、`build_id`が`unknown`になります。
テストはMicrosoft.Testing.Platformで実行します（[global.json](../global.json)）。直接実行する時は`dotnet test --solution xlwright.slnx`です。

実Excelを起動するテスト（`tests/XlWright.RealExcel.Tests`）は、`xlwright.slnx`に入れていないため、確認の入口とcommitのhookでは動きません。専用のExcelを起動して終了させます。日常のExcelには接続せず、実行の前からあるExcelのプロセスが変わらないことと、テストが起動したExcelが残らないことを確かめます。

```powershell
dotnet test --project tests\XlWright.RealExcel.Tests
```

Markdownのチェック:

```powershell
npm run lint:md
```

commitのhookと同じチェックを、全件に対して行う:

```powershell
prek run --all-files
```

hookの設定は [prek.toml](../prek.toml)、Markdownのルールは [.markdownlint-cli2.jsonc](../.markdownlint-cli2.jsonc) にあります。

## 開発版のプラグイン

確認の入口が最後に通したビルドを、Claude Codeの会話から使うためのプラグインです。設計は[サーバーとビルドとプラグイン](architecture.md#サーバーとビルドとプラグイン)にあります。
ユーザーの設定`~/.claude/settings.json`へ書き込むため、入れるのは人です（または人の承認を得て行います）。確認の入口を1回通してから、`env`の`CLAUDE_CODE_PLUGIN_DIRS`に`~/<リポジトリまでのパス>/plugins/xlwright-dev`を加えます。ほかの値がある時は、`;`で区切って後ろに足します。

```json
{
  "env": {
    "CLAUDE_CODE_PLUGIN_DIRS": "~/projects/xlwright/plugins/xlwright-dev"
  }
}
```

新しく始めた会話に、`mcp__plugin_xlwright-dev_xlwright__server_info`が現れます。開いたままの会話は、始めた時のビルドを使い続けます（[記録](notes/dev-plugin-loading.md)）。
この設定は、すべてのプロジェクトの会話で開発版のサーバーを起動します。開発版のサーバーには`--project`としてこのリポジトリの完全なパスが渡るため、ほかのプロジェクトのアプリでは開始できません。
戻すには、加えたパスを消します。
