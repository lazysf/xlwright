# 利用する版の配り方

利用する版のプラグイン`xlwright`の配り方を、バイナリをプラグインに同梱しない形で比べた記録です。人の判断で、[設計](../architecture.md#サーバーとビルドとプラグイン)の「自己完結の形で発行し、プラグインに入れる」を見直しています（2026-11-05）。第1候補は、NuGetと`dnx`です。設計はまだ変えていません。

## 決めたことと保留したこと

- 決めたこと（人、2026-11-05）: バイナリをプラグインに同梱しない。.NET 10 SDKは、利用者がそれぞれ入れる。
- 保留したこと（人、2026-11-05）: サーバーの取得元。第1候補は、nuget.orgと`dnx`。要件の「.NETを入れる手順を省く」と設計の「自己完結で同梱する」は、取得元を決める時に書き直す。それまで、配布に関わる作業（利用する版の発行、会社のPCでの確認、TFMの`net10.0`への変更）は行わない。

## 版

- Windows 11 Pro 10.0.26300、.NET SDK 10.0.402と10.0.403
- Claude Codeの版は、このPCのシェルに`claude`が無く、確かめていない

## 一般のMCPサーバーの配り方

文書と先行事例から分かったことです。実機では確かめていません。

- Node: 利用者がNodeを入れ、`npx -y <パッケージ>@<版>`でnpmから取って起動する。最も多い形。
- Python: 利用者がuvを入れ、`uvx <パッケージ>`でPyPIから取って起動する。
- .NET: 利用者が.NET 10 SDKを入れ、サーバーをNuGetのツールのパッケージ（`PackageType`が`McpServer`）として出し、`dnx <id>@<版> --yes`で起動する。`dnx`は.NET 10 SDKに付く（[Microsoftの手引き](https://learn.microsoft.com/en-us/dotnet/ai/quickstarts/build-mcp-server)）。テンプレートは、OSごとの自己完結のパッケージを既定で作る。
- Docker: 利用者がDockerを入れ、イメージを起動する。WindowsのExcelをCOMで扱うxlwrightには合わない。
- 単体の実行ファイル: GoやRustのサーバーで多い。GitHub Releasesに置き、利用者が取る。
- ExcelMcp: npmの本体`@sbroenne/mcp-server-excel`は5 KBの起動のスクリプトだけで、Node 18以上を求める。実体は、OSごとの`@sbroenne/mcp-server-excel-win32-x64`（161 MB、4ファイル）にある。npmを、バイナリを配る場所として使っている。
- Claude Codeのプラグインの入手元には`archive`（HTTPSのzipを`sha256`で固定する）がある。1ファイル512 MiB、合計1 GiBまで。v2.1.224以降が要る（[マーケットプレイスの公開](https://code.claude.com/docs/en/plugins/host-marketplace.md)）。同梱の形で配る時の候補だった。

## `dnx`の実験

### 再現の手順

1. `src/XlWright.Core`と`src/XlWright.McpServer`と、リポジトリの直下の`Directory.Build.props`・`Directory.Packages.props`・`global.json`を、セッションの一時ディレクトリへ写します。リポジトリは変えません。
2. 写しの`XlWright.McpServer.csproj`に、`PackAsTool`・`ToolCommandName`（`xlwright`）・`PackageId`（`XlWright.Probe`）・`Version`（`0.0.1-probe`）・`PackageType`（`McpServer`）・`RuntimeIdentifiers`（`win-x64`）を書き、`dotnet pack -c Release`で一時ディレクトリのフォルダへ出します。
3. `NUGET_PACKAGES`と`DOTNET_CLI_HOME`を一時ディレクトリに向け、`dotnet dnx XlWright.Probe@0.0.1-probe --yes --source <フォルダ>`を2回実行します。
4. 空の`DOTNET_CLI_HOME`で、`-- --project <リポジトリ> --no-excel`を付けて起動し、`initialize`の要求を標準入力へ送ります。標準出力と標準エラーを分けて受けます。

### 判定

- `net10.0-windows`のままの`PackAsTool`: 不一致。`NETSDK1146`（`PackAsTool`は`TargetPlatformIdentifier`の付いたTFMに対応しない）で止まった。
- 写しのTFMを`net10.0`にした`PackAsTool`: 支持。指すだけのパッケージ（2 KB）と、`win-x64`のパッケージ（1.6 MB）ができた。自己完結ではなく、.NETの実行環境を使う形だった。警告は、Windowsだけの API を呼ぶ4か所の`CA1416`（`Thread.SetApartmentState`・`Type.GetTypeFromProgID`・`Marshal.FinalReleaseComObject`の2か所）だけだった。
- `dnx`での起動: 支持。2回とも、約1秒で使い方を出して終了コード2で終わった。パッケージは、`<NUGET_PACKAGES>/xlwright.probe/0.0.1-probe/`と`<NUGET_PACKAGES>/xlwright.probe.win-x64/0.0.1-probe/`に、版ごとのフォルダとして置かれた。
- 初回の.NETの案内文: 標準エラーに出た。標準出力は`initialize`の応答の1行だけで、MCPの通信は壊れなかった。

## 形の比較

xlwrightの要件（[提供形態](../requirements.md#利用環境と提供形態)）に照らした比較です。どれもバイナリをプラグインに同梱しません。

- NuGetと`dnx`: 利用者は.NET 10 SDKを入れる。プラグインの`.mcp.json`に`dnx`と版を書く。版ごとのフォルダに置かれるため、使用中の会話を止めずに新しい版を出せる見込み（推測。古い版のフォルダを使う会話が残っても、新しい版は別のフォルダに入るため）。TFMを`net10.0`へ変え、`SupportedOSPlatform`で`CA1416`を扱う変更が要る。nuget.orgへの公開は、人のアカウントとAPIキーで行う。
- `dotnet tool install -g`: 同じパッケージを、利用者が一度だけ入れる。更新は`dotnet tool update`で、動いているサーバーが実行ファイルを掴むと失敗しうる。使用中の会話を止めずに更新する要件に合いにくい。
- framework-dependentのzipを置き、利用者が.NETの実行環境を入れる: zipを取る経路（プラグインへの同梱か、自前の取得）が別に要る。
- npmの包み（ExcelMcpの形）: 利用者はNodeも入れる。npmに、OSごとのバイナリを出す。
- MCPBの束（`.mcpb`）: Claude Codeのプラグインは読めるが、バイナリを束に入れるため同梱の形と同じ。

## GitHubから取る形

- GitHubのリポジトリ`lazysf/xlwright`は、公開の設定で作られている。中身は空で、まだpushしていない（GitHubのAPIで確認、2026-11-05）。最初のpushで、履歴がすべて公開される。
- プラグイン（`.mcp.json`とスキル）は、リポジトリに`.claude-plugin/marketplace.json`を置けば、GitHubのマーケットプレイスとして配れる。バイナリは要らない。
- サーバーの実体をGitHubだけから取る素直な方法は無い。
  - GitHub PackagesのNuGetは、公開のパッケージでも取る時にアクセストークンが要る（推測。文書では確かめていない）。
  - `dnx`が読むのはNuGetのフィードかローカルのフォルダで、GitHub Releasesのzipは読めない。取りに行くスクリプトを自前で書くことになる。
  - ソースをプラグインに入れて利用者のSDKでビルドする形でも、依存のパッケージ（`ModelContextProtocol`など）をnuget.orgから取る。初回の起動は数十秒かかり、MCPの起動の期限を越えうる（推測）。
- このため、どの形でもnuget.orgに届くことが前提になる見込み。

## 未確認

- nuget.orgへの公開（アカウント、APIキー、パッケージの名前の予約）。人の操作が要る。
- 会社のPCで、.NET 10 SDKを入れられるか。nuget.orgに届くか（プロキシや管理設定）。段階4の未確認（[設計](../architecture.md#サーバーとビルドとプラグイン)）と一緒に確かめる。
- Claude Codeのプラグインの`.mcp.json`の`command`に`dnx`を書いた時に、起動できるか（`dnx`は`dotnet`の隣の`dnx.cmd`）。
- TFMを`net10.0`へ変えた時の、本体・テスト・確認の入口への影響。
- 版を上げた時に、開いたままの会話が古い版を使い続けるか。

## 後始末

- 一時ディレクトリの写し、パッケージ、NuGetのキャッシュを消した。
- リポジトリは変えていない（`git status`でこの記録だけ）。
