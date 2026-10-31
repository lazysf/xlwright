# xlwright

AIエージェントによるExcel業務アプリケーションの開発・実行・検証を支援する基盤です。
Windows上の実Excelを操作し、VBAソースの反映、マクロ実行、テスト、結果取得の反復を自動化することを目指します。

## 現在の状態

リポジトリとしての要件・設計・検証計画は [docs/](docs/README.md) で管理しています。
実機検証の準備段階R1は完了しました。R2（基盤のSpike）は進行中で、T7の起動設定評価は完了しました。T8のセル入出力評価は完了しました。T9の計算結果の鮮度評価と設計採用も完了しました。T10以降は未着手です。
MCP接続基盤と開発用ツールを実装しています。製品のExcel操作APIとPoCは未実装です。Spike-0の初回コンソールと検証用資産を準備しています。
ビルドと非COM自己チェック、ユーザーによる初回正常系と次回開始を確認済みです。
直接のエージェントシェルではCOM生成のアクセス拒否を観測しました。アプリのMCP経由では、固定コンソールによる正常系と次回開始を確認しました。T5は2026-10-30にユーザーの完了承認を得ました。[R1の完了判定と未確認範囲](docs/notes/spike-0-results.md#r1の完了判定と後続の確認)を参照してください。製品サーバー内のSTA構成、L2・L3全体、S13全体は未合格です。
技術的な成立性や実運用の安全性は、まだ確認されていません。

開発の順序は [ROADMAP.md](ROADMAP.md)、直近の作業は [TODO.md](TODO.md) で管理しています。
具体的な実験と判定基準は [Spike-0](docs/spike-0.md) を参照してください。
P0-coreの完成前に、小さなブックでソース編集・自動反映・実行・結果確認の反復を成立させます。
初期の異常時は、人による確認・後始末とサーバー再起動を許容します。

## 開発環境

- Windows 11 x64
- Windows PowerShell 5.1（OS標準）
- .NET 10 SDK（選択規則は [global.json](global.json)）
- Git
- prek（動作確認済み: 0.5.4）
- nvm-windows（動作確認済み: 2.0.1）とNode.js（版は [.nvmrc](.nvmrc)）

Windows向けのスクリプトとコマンドはWindows PowerShell 5.1を前提とします。
PowerShell 7や`pwsh`の追加導入は必要条件にしません。
Codex内蔵シェルのPowerShellの版は、ユーザーの通常シェルの版を示しません。

Excelを使う実機検証の対象は Microsoft 365 Excel 64-bit のデスクトップ版です。
Excelの設定変更や実機検証は、初期セットアップには含みません。

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

構成と作業手順は[開発ガイド](docs/development.md)に記載しています。
AIエージェント向けのリポジトリ固有ルールは [AGENTS.md](AGENTS.md) にあります。

## 参考資料

[初期のプロジェクト概要 V22](references/PROJECT_OVERVIEW_V22.md) は、正式文書を整備する出発点として取り込んだ参考資料です。
今後の仕様判断には、現在の `docs/` と実装・検証結果を使います。

## 開発時のチェック

commit 時に、行末空白、末尾改行、大きな追加ファイル、Markdownの記法をチェックします。
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
