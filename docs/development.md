# 開発ガイド

本書は、xlwrightの開発の手順と確認の入口を記載します。
進め方の規則は [AGENTS.md](../AGENTS.md#開発の進め方) を正本とします。

## 進め方

- 作業は周期で進めます。周期の作るものと完了の条件は [TODO.md](../TODO.md) に書きます。
- Excelに触れない振る舞いは、テストを先に書いて失敗を確かめ、実装して通し、整えます。
- 実Excelの振る舞いは、捨てる前提の実験で確かめ、分かった振る舞いを結合テストで固定します。実験は`spikes/`に置きます。
- テストは、Excelを起動しないものと、実Excelを起動するものを分けます（[設計](architecture.md#本体の部品の境界)）。

## 確認の入口

今ある入口は次のとおりです。ビルドとテストの入口は、段階0で作る時に書き足します。

Markdownのチェック:

```powershell
npm run lint:md
```

commitのhookと同じチェックを、全件に対して行う:

```powershell
prek run --all-files
```

hookの設定は [prek.toml](../prek.toml)、Markdownのルールは [.markdownlint-cli2.jsonc](../.markdownlint-cli2.jsonc) にあります。
