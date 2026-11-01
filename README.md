# xlwright

AIエージェントによるExcel業務アプリケーションの開発・実行・検証を支援する基盤です。
Windows上の実Excelを操作し、VBAソースの反映、マクロ実行、テスト、結果取得の反復を自動化することを目指します。

## 現在の状態

R3（PoC-1）が進行中です。MCP接続基盤と、Spike0から移したCoreの一部を実装しています。製品のExcel操作APIとPoCは未実装です。
技術的な成立性や実運用の安全性は、まだ確認されていません。

要件と設計は [docs/](docs/README.md)、開発の順序は [ROADMAP.md](ROADMAP.md)、直近の作業は [TODO.md](TODO.md) で管理しています。
実験と判定基準は [Spike-0](docs/spike-0.md) を参照してください。

## 開発環境

- Windows 11 x64
- Windows PowerShell 5.1（OS標準）
- .NET 10 SDK（選択規則は [global.json](global.json)）
- Git
- prek（動作確認済み: 0.5.4）
- nvm-windows（動作確認済み: 2.0.1）とNode.js（版は [.nvmrc](.nvmrc)）

Windows向けのスクリプトとコマンドはWindows PowerShell 5.1を前提とします。

Excelを使う実機検証の対象は Microsoft 365 Excel 64-bit のデスクトップ版です。
Excelの設定変更や実機検証は、初期セットアップには含みません。
Spike0のビルドには、Officeの既存PIAが必要です。

リポジトリのルートで、SDKの選択とGit hookの登録を確認してください。

```powershell
dotnet --version
prek install
```

Markdownのチェックに使うNode.jsは、nvm-windowsで `.nvmrc` の版を選択します。
その後、[package-lock.json](package-lock.json) に固定した開発用パッケージをインストールします。

```powershell
nvm install (Get-Content .nvmrc)
nvm use (Get-Content .nvmrc)
npm ci
```

開発用パッケージにはMarkdownのチェックと、MCP Inspector 2.13.0を含みます。

構成と作業手順は[開発ガイド](docs/development.md)に記載しています。
AIエージェント向けのリポジトリ固有ルールは [AGENTS.md](AGENTS.md) にあります。

## 参考資料

[初期のプロジェクト概要 V22](references/PROJECT_OVERVIEW_V22.md) は、正式文書を整備する出発点として取り込んだ参考資料です。
今後の仕様判断には、現在の `docs/` と実装・検証結果を使います。

## 開発時のチェック

通常のビルド・非COM検証・Inspector CLIによる接続検証は、次のコマンドでまとめて実行します。
実Excelは起動しません。

```powershell
npm run check
```

チェック後に`npm run inspect`でInspector CLIからツール一覧を取得できます。
GUIを使う場合だけ`npm run inspect:web`を実行します。
操作方法は[開発ガイド](docs/development.md#inspectorによる調査)を参照してください。
ツール定義のスナップショットと、実Excel実験の固定コピーの作成も同じ文書に記載しています。

commit 時に、行末空白、末尾改行、大きな追加ファイル、Markdownの記法をチェックします。
Markdown 以外のファイルの追加・変更を含む commit では `npm run check` も実行します。
全件チェックは次のコマンドで実行します。

```powershell
prek run --all-files
```

設定は [prek.toml](prek.toml) にあります。
取り込んだ参考文書 `references/PROJECT_OVERVIEW_V22.md` は、自動整形とMarkdownのチェックの対象から除外しています。

Markdownのチェックはmarkdownlint-cli2で行います。
`npm ci` の後、Markdownだけを全件チェックする場合は次のコマンドを実行します。

```powershell
npm run lint:md
```

ルールの設定は [.markdownlint-cli2.jsonc](.markdownlint-cli2.jsonc) にあります。

## ライセンス

[MIT License](LICENSE)
