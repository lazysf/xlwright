# TODO

今の周期の作業だけを書きます。段階と完了の根拠は [ROADMAP.md](ROADMAP.md) にあります。
周期の決め方は [AGENTS.md](AGENTS.md#開発の進め方) にあります。

## 周期1（段階0）: 空のサーバーと確認の入口

作るもの: `server_info`だけを持つ空のサーバーと、AIエージェントが人なしで回す確認の入口。

完了の条件:

- 確認の入口の1つのコマンドで、ビルド、Excelを起動しないテスト、Inspectorでの`server_info`、Markdownのチェックが通る。
- `server_info`の応答が、[API](docs/api.md#server_info段階0)の契約と一致する。

作業:

- [x] 使う依存（.NETのSDKの版、MCPのSDK、テストのフレームワーク）の名前と版の一覧を作り、周期の始めに人の承認を得る。新しい依存を加えるためです。
- [x] ExcelMcpの`McpServer`の構成と、テストの作り（`tests/`）を読む。
- [x] `server_info`の振る舞いのテストを先に書き、失敗を確かめる。
- [x] `src/XlWright.Core`・`src/XlWright.McpServer`を作り、テストを通す。
- [x] 確認の入口を作り、[開発ガイド](docs/development.md)に書き足す（[サーバーとビルドとプラグイン](docs/architecture.md#サーバーとビルドとプラグイン)）。
- [x] commitのhookに、確認の入口を戻す。

## 書き留めたこと

周期の中で後へ回した事項を書きます。

- ROADMAPの段階1の作るものに`range_write`があるが、到達する状態にセルへの書込みが無い。段階1に入る時に、どちらかに揃える。
- 古いビルドを消す工程は、段階0の未確認（使用中のビルドのフォルダの名前の変更が失敗するか）を確かめてから作る。
- `plugins/xlwright-dev/.mcp.json`を`.gitignore`に加えるのは、開発版のプラグインを作る周期で行う。
- 設計の「終了」の未確認の後半（人が閉じた後もプロセスが残るか）は、段階2の「人への引継ぎ」の未確認と重なる。段階1に入る時に整理する。
- OLEのメッセージフィルターが実Excelで働くかを、ROADMAPの段階1の確かめることに明示する。
- `server_info`の失敗の本文は`detail`を省いている。APIの「成功と失敗」に、ツールの契約が`detail`を定めない時の扱いを書くかを、人が判断する。
- 確認の入口を通らないビルドの`build_id`は`unknown`で、APIの形の外にある。APIに書くかを、人が判断する。テストの正規表現は`unknown`を通すため、`-p:XlWrightBuildId`が渡らない回帰はInspectorの工程だけが検出する。
- `scripts/check.ps1`はBOMの無いUTF-8で、Windows PowerShell 5.1はCP932として読む。今は日本語がコメントにしか無い。ASCIIだけで書くか、BOMを付けるかを決める。
- commitのhookの中のビルドは、stagedの変更があるため、常に`<親のcommit>-dirty-…`になる。開発版のプラグインで新しいビルドを確かめる周期で、扱いを決める。
- 確認の入口のInspectorの呼出しと、テストの`McpClient.CreateAsync`に期限が無い。サーバーが固まると止まったままになる。
- `git diff --check HEAD`は、追跡していないファイルを見ない。
- `ServerOptions`は、`--project --no-excel`の`--no-excel`をパスとして受け取り、`--project`の2回目で黙って上書きする。段階1のパスの検査で扱う。
- `server_info`の引数の拒否は、すべてのキーを`message`に並べるため、応答の20,000文字の上限を超えうる。
- リポジトリのパスに空白がある時に、InspectorのCLI が`--project`を1つの引数として渡すかは、確かめていない。
