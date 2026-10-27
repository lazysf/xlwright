# xlwright

現在はリポジトリの初期構成のみを用意しています。

## 開発時のチェック

動作確認済みのバージョンは prek 0.5.4 です。clone 後に Git hook を登録してください。

```powershell
prek install
```

commit 時に、行末空白、末尾改行、大きな追加ファイルをチェックします。
全件チェックは次のコマンドで実行します。

```powershell
prek run --all-files
```

設定は [prek.toml](prek.toml) にあります。
設計ベースライン文書 `references/PROJECT_OVERVIEW_V22.md` は、自動整形の対象から除外しています。

## ライセンス

[MIT License](LICENSE)
