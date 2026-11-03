# 使用中のビルドのフォルダの名前の変更

[設計](../architecture.md#サーバーとビルドとプラグイン)の段階0の未確認「開いた会話が使っているビルドのフォルダは、名前の変更が失敗するか」の前半を確かめた記録です。

## 版

- Windows 11 Pro 10.0.26300、Windows PowerShell 5.1.26100、.NET SDK 10.0.403
- ビルド: `8b256f5-dirty-20261103T051630Z`（確認の入口が通したもの）

## 再現の手順

```powershell
powershell -NoProfile -File spikes\build-folder-rename\Test-BuildFolderRename.ps1
```

[スクリプト](../../spikes/build-folder-rename/Test-BuildFolderRename.ps1)は、`latest.json`が指すビルドを`.xlwright/spikes/`の下へ写し、写しからサーバーを`--no-excel`で起動します。動いている間と、標準入力を閉じてサーバーが終わった後に、写しのフォルダの名前の変更を試します。

## 判定

支持。

- サーバーが動いている間: `Rename-Item`が`IOException`（アクセスの拒否）で失敗した。
- サーバーが終わった後: 名前の変更が成功した。

Claude Codeのデスクトップアプリの会話が、開発版のプラグインから起動したサーバーでも、同じ結果だった。2つの会話が開いたまま、それぞれが使うビルドのフォルダ（`25c4f69-dirty-20261103T053258Z`と`25c4f69-dirty-20261103T053833Z`）で`Rename-Item`を試し、どちらも`IOException`（アクセスの拒否）で失敗した。フォルダは元の名前のまま残った。

## 後始末とGit差分

- 写しのフォルダ（`.xlwright/spikes/build-folder-rename/<時刻>/`）を消した。`.xlwright/`はGit管理外です。
- Gitの差分は、スクリプトと本書の追加だけです。
