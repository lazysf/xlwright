# 開発版のプラグインの読み込み

[設計](../architecture.md)の段階0の未確認のうち、開発版のプラグインに関わる3つを確かめた記録です。会話ごとのサーバーの起動、`${CLAUDE_PROJECT_DIR}`の展開、新しい会話だけで新しい版が読まれるか、の3つです。

## 版

- Windows 11 Pro 10.0.26300、Claude Codeのデスクトップアプリ（Codeタブ）、Windows PowerShell 5.1.26100
- ビルド: `25c4f69-dirty-20261103T051947Z`から`25c4f69-dirty-20261103T054116Z`まで、確認の入口が通した4つ

## 再現の手順

1. 確認の入口（`scripts\check.ps1`）を通し、`plugins/xlwright-dev/.mcp.json`を書き出します。
2. 開発版のプラグインを読み込ませます。下の「入れ方」のとおり、3つの経路を試しました。
3. 新しい会話Aで、`mcp__plugin_xlwright-dev_xlwright__server_info`を呼びます。Aを開いたまま確認の入口を通し、新しい会話Bと、開いたままのAで、もう一度呼びます。
4. `Win32_Process`で、`XlWright.McpServer.dll`を起動している`dotnet`のコマンドラインと親のプロセスを照会します。

## 判定

### 入れ方

- ローカルのフォルダのマーケットプレイス: 不一致。デスクトップアプリのマーケットプレイスの追加の画面は、`https://`のURLかGitHubの`owner/repo`しか受け付けなかった。チャットで`/plugin`を打っても、同じ画面が開いた。
- プロジェクトの`.claude/settings.local.json`の`env`に`CLAUDE_CODE_PLUGIN_DIRS`を書く形: 不一致。新しい会話にプラグインは読み込まれず、サーバーのプロセスも無かった。
- ユーザーの設定`~/.claude/settings.json`の`env`の`CLAUDE_CODE_PLUGIN_DIRS`に、`plugins/xlwright-dev`を加える形: 支持。新しい会話に、`mcp__plugin_xlwright-dev_xlwright__server_info`が現れた。

設定を変えた後に、アプリは再起動していない見込みです。会話のプロセスの親の`claude.exe`の生成時刻が、設定を変えた時刻より前でした。人には確認していません。

### 新しい会話だけで新しい版が読まれるか

支持。

- 会話Aの1回目: `…T053258Z`（その時点の最新）
- 確認の入口で`…T053833Z`を作った後の会話B: `…T053833Z`
- 開いたままの会話Aの2回目: `…T053258Z`（前の版のまま）

### 会話ごとのサーバーの起動

支持。2つの会話が同じビルドを使った時、サーバーは別々の`claude.exe`の子として、2つ起動していた。

### `${CLAUDE_PROJECT_DIR}`の展開

- `${CLAUDE_PROJECT_DIR:-.}`: 不一致。コマンドラインは`--project .`で、既定の値に落ちた。
- `${CLAUDE_PROJECT_DIR}`: 支持。コマンドラインは`--project <リポジトリの完全なパス>`だった。

[MCPの文書](https://code.claude.com/docs/en/mcp.md)は、プラグインの設定では`${CLAUDE_PROJECT_DIR}`を直接置き換え、既定の値は要らないと書いています。既定の値の付いた形が置き換えの対象から外れることは、文書に書かれていません。

## 後始末とGit差分

- プロジェクトの`.claude/settings.local.json`は、効かなかったため消した。
- ユーザーの設定の`CLAUDE_CODE_PLUGIN_DIRS`の追加は、人の承認を得て行い、残している。戻すには、追加した`;~/projects/xlwright/plugins/xlwright-dev`を消す。
- 動いている会話のビルドのフォルダの名前の変更も、このとき試した（[記録](build-folder-rename.md)）。
- Gitの差分は、確認の入口の`.mcp.json`の書き出し、プラグインのマニフェスト、`.gitignore`、文書の更新です。
