# xlwright

AIエージェントによるExcel業務アプリケーションの開発・実行・検証を支援する基盤です。
Windows上の実Excelを操作し、VBAソースの反映、マクロ実行、テスト、結果取得の反復を自動化することを目指します。

## 現在の状態

リポジトリを片付け、開発をしなおしているところです。コードはまだありません。次は[ロードマップ](ROADMAP.md)の段階0です。

要件と設計は [docs/](docs/README.md)、開発の順序は [ROADMAP.md](ROADMAP.md)、今の周期の作業は [TODO.md](TODO.md) で管理しています。
AIエージェント向けのリポジトリ固有ルールは [AGENTS.md](AGENTS.md) にあります。

## 開発環境

- Windows 11 x64
- Windows PowerShell 5.1（OS標準）
- Git
- prek
- nvm-windowsとNode.js（版は [.nvmrc](.nvmrc)）

Windows向けのスクリプトとコマンドは、Windows PowerShell 5.1を前提とします。
Excelを使う実機の確認の対象は、Microsoft 365 Excel 64-bit のデスクトップ版です。
.NETの設定は、段階0でコードを作る時に加えます。

リポジトリのルートで、Git hookを登録し、開発用のパッケージを入れます。

```powershell
prek install
nvm install (Get-Content .nvmrc)
nvm use (Get-Content .nvmrc)
npm ci
```

確認の入口は[開発ガイド](docs/development.md)にあります。

## ライセンス

[MIT License](LICENSE)
