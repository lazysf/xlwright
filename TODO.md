# TODO

今の周期の作業だけを書きます。段階と完了の根拠は [ROADMAP.md](ROADMAP.md) にあります。
周期の決め方は [AGENTS.md](AGENTS.md#開発の進め方) にあります。

## 周期9（段階1）: 段階1のツールの契約

作るもの: 段階1の8つのツール（`session_start`・`session_status`・`session_end`・`workbook_save`・`range_read`・`range_write`・`vba_sync`・`macro_run`）の引数・結果・失敗・注釈を[API](docs/api.md#各ツールの契約)に書き、人の承認を得る。実装は承認の後の周期で行う。

完了の条件:

- 8つのツールの契約が、APIの共通の規約と設計（段階1で確かめた事実を含む）と食い違わない。
- 人が契約を承認するか、直す点を示す。

作業:

- [x] 契約をAPIに書く。
- [x] 設計と要件に照らして見直し、確認の入口を通す。
- [x] 人の承認を得る（2026-11-03）。

## 書き留めたこと

周期の中で後へ回した事項を書きます。

- 古いビルドを消す工程は、段階0の未確認（使用中のビルドのフォルダの名前の変更が失敗するか）を確かめてから作る。
- テストの正規表現は`build_id`の`unknown`を通すため、`-p:XlWrightBuildId`が渡らない回帰はInspectorの工程だけが検出する。
- commitのhookの中のビルドは、stagedの変更があるため、常に`<親のcommit>-dirty-…`になる。開発版のプラグインで新しいビルドを確かめる周期で、扱いを決める。
- 確認の入口のInspectorの呼出しと、テストの`McpClient.CreateAsync`に期限が無い。サーバーが固まると止まったままになる。
- `git diff --check HEAD`は、追跡していないファイルを見ない。
- `ServerOptions`は、`--project --no-excel`の`--no-excel`をパスとして受け取り、`--project`の2回目で黙って上書きする。段階1のパスの検査で扱う。
- `server_info`の引数の拒否は、すべてのキーを`message`に並べるため、応答の20,000文字の上限を超えうる。
- リポジトリのパスに空白がある時に、InspectorのCLIが`--project`を1つの引数として渡すかは、確かめていない。
- 設計の段階4の未確認で「成立しない時」の扱いに、ローカルのフォルダのマーケットプレイスを挙げている。デスクトップアプリの画面では追加できなかったため、段階4に入る時に扱いを見直す。
- コマンドラインを読む`NtQueryInformationProcess`の`ProcessCommandLineInformation`は、Microsoftが変わりうるとしているntdllの関数。製品で使うか、WMIのような代わりを使うか（新しい依存になりうる）を、所有の確定を作る周期で決める。
- 段階1の残りの未確認は、製品のサーバーで確かめる2つ（実行の間の`session_status`、日本語を含む応答の文字数とトークン）。
