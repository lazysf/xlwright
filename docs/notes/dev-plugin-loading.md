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

プロジェクトの設定の形は、アプリを再起動した後にも試しました（2026-11-03）。ユーザーの設定から`plugins/xlwright-dev`を外し、`.claude/settings.local.json`の`env.CLAUDE_CODE_PLUGIN_DIRS`にだけ書いてから、人がアプリを再起動しました。最上位の`claude.exe`の生成時刻は、設定を書いた後でした。このリポジトリの新しい会話でも、判定は不一致でした。xlwrightのツールは現れず、`XlWright.McpServer`のプロセスも無く、会話のシェルの`CLAUDE_CODE_PLUGIN_DIRS`はユーザーの設定の値でした。

- 再起動の後の`.claude/settings.local.json`の形は、もう一度試しても不一致でした（2026-11-03、ほかの形の変数を外した後）。
- プロジェクトの共有の設定`.claude/settings.json`の`env`に書く形: 不一致（2026-11-03）。ほかの形の変数を外し、設定を書いた後に人がアプリを再起動しました。このリポジトリの新しい会話にxlwrightのツールは現れず、サーバーのプロセスも無く、会話のシェルの`CLAUDE_CODE_PLUGIN_DIRS`はユーザーの設定の値でした。
- デスクトップアプリのCodeのプロジェクトの環境の`.env`に`CLAUDE_CODE_PLUGIN_DIRS`を書く形: 支持（2026-11-03）。値は`~/.claude/mods/usage-band;~/projects/xlwright/plugins/xlwright-dev`です。ユーザーの設定は`~/.claude/mods/usage-band`だけにし、`.claude/settings.local.json`には`env`を書かず、人がアプリを再起動しました。このリポジトリの新しい会話に`mcp__plugin_xlwright-dev_xlwright__server_info`が現れ、`build_id`は`latest.json`のビルドと一致しました。サーバーは、再起動の後に生成された`claude.exe`の子でした。`~`で始まるパスは、この形でも読み込まれました。
- 会話のシェルの`CLAUDE_CODE_PLUGIN_DIRS`は、`.env`の値の`~`を展開したものでした。ユーザーの設定の値とはつながっていませんでした。
- `.env`の形では、ほかのプロジェクトの新しい会話に、xlwrightのツールは現れず、サーバーのプロセスも増えませんでした。人が新しい会話を開いて確かめました。
- リポジトリの直下の`.env`ファイルに同じ1行を書く形: 不一致（2026-11-03）。アプリの環境から変数を外し、人がアプリを再起動しました。このリポジトリの新しい会話に、xlwrightのツールは現れず、サーバーのプロセスも無く、会話のシェルの`CLAUDE_CODE_PLUGIN_DIRS`はユーザーの設定の値でした。

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
- ユーザーの設定の`CLAUDE_CODE_PLUGIN_DIRS`への追加は、人の承認を得て行い、プロジェクトの環境の`.env`の形を確かめる前に人が外した。今は、デスクトップアプリのプロジェクトの環境の変数で読み込んでいる。
- 再起動の後に試した`.claude/settings.local.json`の`env`、`.claude/settings.json`、リポジトリの直下の`.env`ファイルは、効かなかったため消した。
- 動いている会話のビルドのフォルダの名前の変更も、このとき試した（[記録](build-folder-rename.md)）。
- Gitの差分は、確認の入口の`.mcp.json`の書き出し、プラグインのマニフェスト、`.gitignore`、文書の更新です。
