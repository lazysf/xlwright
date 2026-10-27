# xlwright

AIエージェントによるExcel業務アプリケーションの開発・実行・検証を支援する基盤です。
Windows上の実Excelを操作し、VBAソースの反映、マクロ実行、テスト、結果取得の反復を自動化することを目指します。

## 現在の状態

リポジトリとしての要件・設計・検証計画は [docs/](docs/README.md) で管理しています。
現在は開発環境と作業手順の準備段階です。
本体コード、Spike-0、PoCは未実装・未実施です。
技術的な成立性や実運用の安全性は、まだ確認されていません。

開発の順序は [ROADMAP.md](ROADMAP.md)、直近の作業は [TODO.md](TODO.md) で管理しています。
具体的な実験と判定基準は [Spike-0](docs/spike-0.md) を参照してください。

## 開発環境

- Windows 11 x64
- .NET 10 SDK（選択規則は [global.json](global.json)）
- Git
- prek（動作確認済み: 0.5.4）

Excelを使う実機検証の対象は Microsoft 365 Excel 64-bit のデスクトップ版です。
Excelの設定変更や実機検証は、初期セットアップには含みません。

リポジトリのルートで、SDKの選択とGit hookの登録を確認してください。

```powershell
dotnet --version
prek install
```

構成と作業手順は[開発ガイド](docs/development.md)に記載しています。
AIエージェント向けのリポジトリ固有ルールは [AGENTS.md](AGENTS.md) にあります。

## 参考資料

[初期のプロジェクト概要 V22](references/PROJECT_OVERVIEW_V22.md) は、正式文書を整備する出発点として取り込んだ参考資料です。
今後の仕様判断には、現在の `docs/` と実装・検証結果を使います。

## 開発時のチェック

commit 時に、行末空白、末尾改行、大きな追加ファイルをチェックします。
全件チェックは次のコマンドで実行します。

```powershell
prek run --all-files
```

設定は [prek.toml](prek.toml) にあります。
取り込んだ参考文書 `references/PROJECT_OVERVIEW_V22.md` は、自動整形の対象から除外しています。

## ライセンス

[MIT License](LICENSE)
