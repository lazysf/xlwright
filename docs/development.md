# 開発ガイド

## 設計の参照先

要件と設計の正本は [docsの索引](README.md) から参照してください。
本書はリポジトリでの作業手順を扱います。
API契約と実機検証の仮説は、それぞれの正式文書で管理します。
実施済みの実験の手順と経過は[Spike-0の実機手順](spike-0-runbook.md)にあります。

## セットアップ

必要な環境と初期セットアップは [README.md](../README.md#開発環境) を参照してください。

[global.json](../global.json) は .NET SDK 10.0.402 を基準とします。
同じ10.0.4xx帯の、10.0.402以上のインストール済みパッチ版を選択します。
プレビュー版は選択しません。
選択規則は [Microsoftのglobal.jsonリファレンス](https://learn.microsoft.com/en-us/dotnet/core/tools/global-json) に従います。

現在のC#プロジェクトはSpike0、MCP接続基盤、MCPテストです。製品のExcel操作APIとCoreは未実装です。
Spike0は.NET 10のWindows Desktop共有フレームワークを使います。
Spike0のビルドには、Officeの既存PIAが必要です。固定版は[環境の確認結果](notes/verification-environment.md)を参照してください。
同じ理由で、GitHubの標準ランナーではビルドできないためCIを置いていません。

## 配置方針

現在存在する入口は次のとおりです。

| パス | 用途 |
| --- | --- |
| `AGENTS.md` | AIエージェント向けのリポジトリ固有ルール |
| `docs/README.md` | 現在の正式文書への入口 |
| `references/PROJECT_OVERVIEW_V22.md` | 取り込んだ初期の参考資料 |
| `docs/development.md` | セットアップと開発手順 |
| `docs/spike-0.md` | 次工程の実施計画と結果記録の方法 |
| `docs/spike-0-runbook.md` | 実施済みの実験の手順と経過 |
| `global.json` | .NET SDKの選択 |
| `.editorconfig` | C#と設定ファイルの編集形式 |
| `prek.toml` | commit前のチェック |
| `.nvmrc` | Node.jsの版の選択 |
| `package.json`・`package-lock.json` | MarkdownのチェックとInspectorに使う開発用パッケージの固定 |
| `.markdownlint-cli2.jsonc` | Markdownのチェックのルールと対象 |
| `scripts/` | 統合チェック、Inspector、実験の準備の入口 |

コードと検証資産は、次の位置へ配置します。
空ディレクトリや仮実装を先に追加する必要はありません。

| パス | 用途 |
| --- | --- |
| `spikes/Spike0/` | 実機検証の実験コードと非COM自己チェック |
| `src/XlWright.Core/` | MCPに依存しない中核処理。現在はSpike0から移した実行枠・所有記録・終了安全確認・常駐STA |
| `src/XlWright.McpServer/` | stdio接続基盤と開発モードの`dev_status`。製品のExcel操作APIは未実装 |
| `tests/XlWright.McpServer.Tests/` | サーバーを実プロセスで起動するxUnitテスト、`Core/`のCoreの単体テスト、ツール定義のスナップショット |
| `tests/fixtures/` | Git管理する基準ブック。現在はマクロなしの`lifecycle.xlsx` |

本体と利用プロジェクトの責務は [設計文書](architecture.md) に従います。
本体リポジトリに、利用プロジェクトの `project.json` や実ブックを先に生成しないでください。

## 作業と検証

### 役割分担

エージェントは、コード・テスト・文書の作成と修正、統合チェック、Inspectorの確認、実験の準備、証拠の照合を自走します。
人は、方針の判断、実Excelの操作と復旧の承認、コンソールでの実験の実行、画面の目視確認を担当します。

### 実装の進め方

1. 正式文書の該当箇所と受け入れ条件を確認します。
2. 実機依存の仮説は[Spike-0](spike-0.md)で確認します。
3. その結果に基づいて、対象工程のコードと必要なテストを追加します。
4. 下記の自己修正ループで、完了条件を満たすまで修正します。
5. 差分を確認し、基準フィクスチャや対象外ファイルへの意図しない変更がないか検査します。
6. 実装後の使い方と確認済みの制約を文書へ反映します。

検証コードを含め、依存パッケージは採用時にバージョンを固定します。
Spikeで本体全体や汎用検証基盤を先に作りません。
検証済みの処理は必要な部分をCoreへ移し、再現に必要な小さな実験は残します。一律の再実装は要求しません。
Excelの起動や設定変更を、通常の静的チェックに混ぜないでください。

### 自己修正ループ

エージェントは次の反復を、人の操作なしで実行します。

1. コードとテストを変更します。
2. `npm run check`を実行します。失敗したら、出力された工程のログから原因を特定して修正し、この手順を繰り返します。
3. ツールの公開内容や応答を変更したときは、`npm run inspect`で実際の応答を確認します。正常系に加えて、不正な入力や拒否の応答も確認します。期待と異なれば1へ戻ります。
4. 確認で見つけた問題は、`tests/`のテストへ追加して固定します。

特定のテストだけを再実行するときは、`npm run check -- -TestFilter "*Name*"`を使います。Spike0とテストをビルドし、一致するテストだけを実行します。
ワイルドカードは先頭と末尾に使えます。実行されたテストが0件なら失敗します。
完了の判断には、部分実行ではなく全工程の`npm run check`を使います。

ツールの名前・説明・スキーマを変更したときは、`npm run snapshot:tools`で`tests/XlWright.McpServer.Tests/tools.snapshot.json`を更新します。
モデルはツール定義を指示として読むため、更新後の差分を確認します。
このループは実Excelを操作しません。

### 完了条件

| 変更の種類 | 完了条件 |
| --- | --- |
| 文書だけ | `npm run lint:md`の通過 |
| コード・スクリプト | `npm run check`の全工程の通過 |
| MCPのツール | 上記に加えて、Inspectorでの応答の確認とスナップショットの差分の確認 |
| 実Excelに依存する動作 | 上記に加えて、[実Excel実験](#実excel実験の進め方)の証拠の照合 |

commit時は、Markdown以外のファイルの追加・変更に対してprekのhookが`npm run check`を実行します。
Markdown以外の変更が削除だけのcommitではhookが動かないため、エージェントが実行します。

### テストの層

Excelなしで検証できる部分を増やし、自走ループで扱う範囲を広げます。
R3のCoreは、COMの境界の外をExcelなしで試験できる形で作ります。

| 層 | 手段 | 実行 |
| --- | --- | --- |
| Coreの単体（R3から） | xUnit | `npm run check` |
| MCPのプロトコル | xUnitからサーバーを実プロセスで起動 | `npm run check` |
| Excel入口の検査 | xUnitから合成リポジトリでSpike0の`--preflight`を実行。Excelのコードへ到達しない | `npm run check` |
| 公開内容 | Inspector CLI、スキーマ検査、スナップショット | `npm run check` |
| 実Excel | 固定コピーの実験 | 承認後に人が実行 |

複数の要求にまたがる状態は、接続を維持する`tests/XlWright.McpServer.Tests/`のクライアントで検証します。
Inspector CLIは1要求ごとに接続を閉じるためです。

### 非COM検証の自動実行

ビルドと非COM検証は`npm run check`の1経路だけです。Excelは起動しません。
通常ユーザー側ではNode.js 26.13.0とnpm 11.20.0を確認しました。
sandbox側で指定版が見つからない場合は、正式な権限経路で実行します。

入口は[scripts/development.ps1](../scripts/development.ps1)です。
Windows PowerShell 5.1で2本のcheck.ps1を実行します。
Debug・Releaseの出力と中間ファイルは`.xlwright/development/artifacts/`へ置きます。
続けてInspector CLIで、legacyとmodernの両プロトコル世代のツール一覧と読取り専用の`dev_status`を検証します。
ツール一覧の検証は、スキーマの`--strict`検査とスナップショットの照合を含みます。
最後にMarkdownチェックと`git diff --check`を実行します。
失敗時は後続工程を実行せず、失敗した工程のログの末尾を出力します。

実行結果は`.xlwright/development/runs/<run-id>/result.json`と工程別の標準出力・標準エラーログへ保存します。
ビルドエラーとテストの失敗は標準出力ログに出力されます。
全工程の実行は、開始時に`.xlwright/development/last-passed.json`を消します。
全工程の通過時に、run IDとSpike0のソース・出力一式のハッシュをここへ記録します。実行中にソースが変わった場合は失敗します。
通常はキャッシュだけで依存を復元します。
キャッシュが不足する場合は、承認された取得範囲を確認して`npm run check -- -Restore`を実行します。
これは公式NuGetだけを指定した設定とlock fileで依存を復元します。依存取得の利用条件は[環境記録](notes/verification-environment.md#mcp接続基盤の依存確認)を参照してください。
SDKの初回処理とキャッシュは`.xlwright/mcp-build/`に限定します。
Nodeパッケージの準備には既存の`npm ci`を使います。
check.ps1は同じ出力先を使うため、統合入口から順番に実行してください。

### Inspectorによる調査

`npm run check`の後、リポジトリ直下で次を実行します。

```powershell
npm run inspect
```

Inspector 2.13.0のCLIがツール一覧をJSONで出力して終了します。
サーバーのビルドがソースより古い場合は起動を拒否します。
CLIの標準引数をそのまま渡せます。読取り専用の状態確認は次のとおりです。

```powershell
npm run inspect -- --method tools/call --tool-name dev_status
```

CLIの終了コードは、0が成功、5がツールの`isError`または未知のツール、6が`--strict`のスキーマ違反です。
失敗の種類は、標準エラーの最終行のJSONにも出力されます。

通常の検証はCLIで完結します。
必要なときだけ`npm run inspect:web`でGUIを起動します。
画面でxlwrightへ接続し、Toolsから引数・応答を確認します。
ProtocolでJSON-RPC、Consoleでサーバーの標準エラーを調査できます。
起動と`dev_status`はExcelを操作しません。

[scripts/inspector.mjs](../scripts/inspector.mjs)は公式CLI・GUIへの薄い起動処理です。
標準機能の`--config`による読取り専用設定を使います。
起動ごとにサーバーのビルド一式を`.xlwright/development/inspector/<session-id>/`へコピーします。
CLIのコピーは終了時に削除します。統合チェックのコピーは成功時に削除し、失敗時は調査のために保持します。GUIのコピーは保持します。
Inspectorが参照するDLLと次のビルド先を分けるため、GUIへ接続したまま`npm run check`を実行できます。
コピー中は統合ビルドと排他します。競合した場合は片方の完了を確認してから実行してください。
既存GUIの接続先は起動時のコピーです。新しいビルドを調査するときはGUIを終了し、`npm run inspect:web`を起動し直します。
終了するときはサーバーをDisconnectしてから、起動したターミナルでCtrl+Cを押してください。
`npm ci`などで依存を入れ直すときも、先にInspectorを終了してください。
WindowsではGUIが読み込んだネイティブモジュールの削除を拒否されることを確認しています。

Inspectorの状態とログはセッションごとにリポジトリ内へ置きます。
secret storeはmemoryを指定し、個人のcredential storeを使いません。
GUIは127.0.0.1へbindし、標準の起動ごとの認証を有効にします。
起動URLやProtocol・Consoleにはローカル情報が含まれるため、出力をそのまま共有しないでください。

Inspectorの成功は、AIクライアントのツール選択や一連の操作の確認を含みません。
AIクライアントへの接続は[登録と更新](#登録と更新)を参照してください。
Hotloadによる自動差替えは採用しません。

### MCPツールの実装基準

ツールを追加・変更するときは、次の基準に従います。
根拠と現在の実装との差は[調査記録](notes/mcp-best-practices.md)を参照してください。

- ツール定義は、モデルが読む指示として書きます。説明には用途、副作用、失敗後に再試行できるかを書きます。
- 入力の検証失敗と業務上の失敗は、`isError`のツール実行エラーで返します。文言には次に取る行動を含めます。
- モデルへ伝える失敗は`McpException`で返します。C# SDKの応答に`ArgumentException`の文言は含まれず、`McpException`の文言は含まれることを確認しています。
- 製品ツールの失敗の文言は、[API契約](api.md#応答の形)の1行のJSONとします。
- 製品ツールの引数は、SDKの型変換に任せずツール側で検査し、同じ形の失敗で返します。SDKの既定の失敗文言は詳細を含まないためです。
- 注釈の`ReadOnly`・`Destructive`・`Idempotent`・`OpenWorld`は全ツールで明示します。注釈を安全制御の根拠には数えません。
- 結果は`outputSchema`に沿う`structuredContent`で返し、同じJSONをテキストでも返します。
- 大きな応答になり得るツールには、範囲や項目を絞る引数を設けます。上限超過のエラーには絞り方を含めます。
- 診断は標準エラーかファイルへ出力します。標準出力はプロトコル専用です。
- セル値やVBAの出力など、ブックに由来する文字列は信頼できないデータとして返します。

### 実Excel実験の進め方

Spikeの実Excel実験は、人がコンソールで起動します。承認の条件は[AGENTS.md](../AGENTS.md#自走と承認)に従います。

1. エージェントが、公式情報で分かる範囲を確認します。実機でしか分からない点を、承認事項へ書きます。
2. エージェントが`npm run check`を通し、`npm run prepare:experiment`で固定コピーを作ります。
3. エージェントが、使用するブック、変更・終了する対象、副作用、停止条件を示して承認を得ます。
4. 人がリポジトリ直下の通常のPowerShellで、表示されたコマンドを実行します。
5. エージェントが、終了コード、診断、所有記録、基準ブックのハッシュを照合し、結果を[Spike-0の記録方法](spike-0.md#結果の残し方)に従って残します。

`npm run prepare:experiment`は、最新の全工程の通過と同じソース・ビルドだけをコピーします。
出力先は`.xlwright/development/prepared/<id>/`です。
`manifest.json`に、check run、commit、作業ツリーの差分有無、基準ブック、全ファイルとソースのハッシュを記録します。
通過の記録がない場合、ソースかビルドが記録と異なる場合、統合チェックの実行中は拒否します。

Spike0のExcel入口は、次の検査を通らなければ開始しません。

- 固定コピーの位置から起動され、全ファイルが`manifest.json`と一致すること。診断の`run.prepared`に、DLLのハッシュとcheck runを記録します。
- 共有ガード`.xlwright/mcp-development/active-smoke.json`が存在しないこと。各実行が作成し、正常終了の確認後に除去します。失敗・結果不明では残し、人の確認まで次の開始を止めます。

`--preflight <Excel引数>`は、この検査だけを実行して戻ります。
これらは固定コピー以外のDLLを誤って実行する事故を防ぐ検査です。検査を追加する前に作った固定コピーと、同じユーザーによる改ざんは防げません。
コピーしたDLLを、Excelや作業データの隔離とは扱いません。

失敗・結果不明では再送せず、証拠を保持します。
未完了の所有記録がある場合は、証拠保全と手動復旧の承認を必要とします。
手動復旧は[設計文書](architecture.md#初期の手動復旧)に従います。
ガードがなければ、Spike0は[次回開始の規則](architecture.md#initial-lifecycle)で所有記録を更新します。置き換えた記録は`run.prepared`に残します。ガードを除去できるのは、対象と証拠を確認した後だけです。

## 継続利用するMCP開発環境

stdio接続基盤です。経緯は[ADR 0002](adr/0002-development-mcp-foundation.md)と[ADR 0003](adr/0003-inspector-centered-development.md)を参照してください。
引数なしでは開発用ツールを公開せず、Excel操作を行いません。製品のExcel操作APIは未実装です。
`--development`で起動すると`dev_status`だけを公開します。

`dev_status`は入力なしで、起動識別子とサーバーのSHA-256を返します。Excelを操作しません。

### 登録と更新

AIクライアントへの直接接続は現在登録していません。
このリポジトリはClaude Code・Codexのプロジェクト設定を持ちません。
開発サーバーはInspectorから起動します。
R3の登録方法は[R3の進め方](spike-0.md#r3の進め方)にあります。手順は登録を確認した後にここへ記載します。

## 実行時の生成物

`.gitignore` は .NETの `bin/`・`obj/`・`TestResults/` と、xlwrightの `.xlwright/`・`.xlwright.lock`・`dist/` を除外します。
基準ブックとVBAソースは除外しません。
検証結果を文書に残すときは、ログ全文をそのまま転載せず、必要な観測結果を記録してください。
