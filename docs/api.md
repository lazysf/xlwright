# API

本書は、xlwrightのMCPのツールの名前と共通の規約を定めます。
ツールの名前・説明・引数・応答は、公開インターフェースです。モデルはツールの定義を、指示として読みます。
各ツールの引数と応答は、そのツールを作る段階で、その段階の実機の確認の直後に本書へ書きます。実装の前に人の承認を得ます。
方式は[設計](architecture.md)、実装の基準は[開発ガイド](development.md#mcpツールの実装基準)にあります。

## 共通の規約

### 名前と分け方

ツールは、読むツールと変えるツールに分けます（[ADR 0013](adr/0013-read-and-edit-tools.md)）。

- 名前は`<対象>_<動詞>`です。
- `<対象>_get`は読むだけです。`<対象>_edit`は変えます。変える操作は、引数`action`で選びます。
- 状態を変える操作、長く走る操作、頻繁に使う操作は、固有の動詞を持つ単独のツールにします。
- 範囲の中身（値・数式）は`range_read`と`range_write`で扱います。それ以外（書式、結合、入力規則、条件付き書式、ロック、コピー、並べ替え、オートフィルター）は`range_get`と`range_edit`で扱います。
- 1回の呼出しは1つの操作です。複数の操作をまとめて送るツールは作りません。部分的に失敗した時の報告を単純にするためです。

### 引数

- 引数はツールの側で検査します。未知の引数は拒否します。
- シートと範囲は、名前とA1形式で明示します。アクティブなブックやシートに依存しません。
- 値と数式は、1つのセルでも2次元の配列で受け渡します。

### 成功と失敗

- 成功は、`outputSchema`に沿う`structuredContent`と、同じJSONの本文で返します。
- 失敗は、`isError: true`の`CallToolResult`で返します。`structuredContent`は付けません（[ADR 0007](adr/0007-error-as-direct-result.md)）。
- 失敗の本文は、次の形の1行のJSONです。

```json
{"error":{"code":"CONFLICT","detail":"SESSION_IN_USE","message":"…","data":{}},"session":{"state":"closed","busy":false}}
```

- `message`には、次に取る行動を書きます。
- `data`は、ツールごとの補足です。`STATE_UNKNOWN`以外では省略できます。部分的に変更した時の処理済みの部分、止まって知らせる時の引継ぎの情報（記録したPIDと生成時刻、所有記録のパス、人が取る手順）、診断の記録のパスを入れます。中身は各ツールの契約で定めます。
- `code`は次の6つです。`detail`の値は、各ツールの契約で定めます。

| `code` | 意味 | セッション |
| --- | --- | --- |
| `INVALID_REQUEST` | 引数の不正、未対応の操作 | 変えない |
| `NOT_FOUND` | 指定したシート・範囲・モジュール・マクロが無い | 変えない |
| `CONFLICT` | 使用中、実行中、状態が合わない | 変えない |
| `EXCEL_ERROR` | Excelが明確なエラーを返した | 変えない |
| `STATE_UNKNOWN` | 結果が不明。`data.excel_ended`（専用のExcelを終えたか）と`data.restartable`（このサーバーで開始し直せるか）の真偽値を必ず付ける | `closed` |
| `INTERNAL_ERROR` | xlwright自身の失敗 | 変えない |

- 「セッション」の列は、既定の扱いです。[終了](architecture.md#終了)と[テスト](architecture.md#テスト)の手順で止まった時は、ほかのコードでも`closed`になります。失敗の本文の`session`が、実際の状態を示します。
- AIエージェントは、`STATE_UNKNOWN`で`excel_ended`と`restartable`がともに真の時だけ、開始し直します。それ以外の時は、`message`の手順に従います。

- 結果不明の扱いは[設計](architecture.md#状態と結果不明)のとおりです。xlwright自身の失敗で専用のExcelの状態が分からなくなった時は、`INTERNAL_ERROR`ではなく`STATE_UNKNOWN`を返します。
- 応答の本文は20,000文字までです。
- 読むツールは、超える時に切り詰めずに`INVALID_REQUEST`で拒否し、範囲の絞り方を`message`に書きます。何も変えていないため、絞って呼び直せます。
- 変える操作と実行（`vba_sync`・`macro_run`・`test_run`など）は、大きさを理由に失敗にしません。拒否すると、AIエージェントが何も起きなかったと読み、呼び直して副作用を2回起こすためです。結果を要約して返し、全体を`<アプリ>/.xlwright/results/<id>.json`へ書き、そのパスを返します。診断の記録には書きません。要件の安全性の「セル値やVBA全文を無条件にログへ出しません」に従うためです。

### 注釈

- 全ツールで`ReadOnly`・`Destructive`・`Idempotent`・`OpenWorld`を明示します。
- `<対象>_get`、`server_info`、`session_status`、`range_read`は、読み取りだけです。
- 注釈を、安全の制御の根拠には数えません。

## ツールの一覧

合わせて33個です。段階は[ロードマップ](../ROADMAP.md)のものです。

| 段階 | ツール | 用途 |
| --- | --- | --- |
| 段階0 | `server_info` | 版とビルドの識別子 |
| 段階1 | `session_start` | アプリのフォルダを渡してセッションを始める |
| 段階1 | `session_status` | 状態、実行中の操作と経過時間 |
| 段階1 | `session_end` | 保存するか（必須の真偽値）を指定して終える |
| 段階1 | `workbook_save` | 主ブックを保存する |
| 段階1 | `range_read` | 値・数式を読む。段階5で表示される文字列を加える |
| 段階1 | `range_write` | 値・数式を書く、中身を消す |
| 段階1 | `vba_sync` | `src/`を主ブックへ反映する。段階3で、シートとブックのモジュールに広げる |
| 段階1 | `macro_run` | 反映してからマクロを実行し、結果を返す |
| 段階3 | `vba_get` | ブックのモジュールの一覧。段階4で、参照設定の一覧を加える |
| 段階3 | `vba_check` | コンパイルの確認 |
| 段階3 | `test_run` | テストの写しでテストを実行し、結果を返す |
| 段階4 | `vba_extract` | 既存のブックから`src/`へ取り出す |
| 段階4 | `vba_edit` | 参照設定の追加と解除 |
| 段階4 | `workbook_create` | 主ブックの無いセッションで、アプリのフォルダの直下に主ブックを作る |
| 段階4 | `workbook_get`・`workbook_edit` | `.xlsm`への変更。段階5で、ブックの保護を加える |
| 段階5 | `excel_get`・`excel_edit` | 計算の方法、再計算 |
| 段階5 | `sheet_get`・`sheet_edit` | シートの一覧と操作、ウィンドウの表示の設定 |
| 段階5 | `range_get`・`range_edit` | 範囲の中身以外 |
| 段階5 | `rows_columns_get`・`rows_columns_edit` | 行・列 |
| 段階5 | `name_get`・`name_edit` | 名前 |
| 段階5 | `table_get`・`table_edit` | テーブル |
| 段階5 | `button_get`・`button_edit` | シートのボタンとマクロの割り当て |
| 段階5 | `page_setup_get`・`page_setup_edit` | 印刷の設定 |

- 段階4で、`macro_run`と`test_run`に、ダイアログへの事前の応答を加えます。
- 段階3で、`session_status`に実行中のテストの情報を加えます。

### 範囲の表との対応

[範囲の表](requirements.md#操作)の各行の操作を、次のツールが受け持ちます。設定できるものの読み取りは、同じ対象の`_get`が受け持ちます。

| 対象 | ツール |
| --- | --- |
| Excel全体 | `excel_get`・`excel_edit` |
| ブック | 新規作成は`workbook_create`、保存は`workbook_save`、`.xlsm`への変更とブックの保護は`workbook_edit` |
| シート | `sheet_get`・`sheet_edit`。ほかのブックからのコピーも`sheet_edit` |
| ウィンドウ | `sheet_get`・`sheet_edit`。表示の設定とウィンドウ枠の固定は、シートごとのウィンドウの設定として扱う |
| セル範囲 | 値・数式の読み書きと消去、計算結果と表示される文字列は`range_read`・`range_write`。使われている範囲は`sheet_get`。それ以外は`range_get`・`range_edit` |
| 行・列 | `rows_columns_get`・`rows_columns_edit` |
| 名前 | `name_get`・`name_edit` |
| テーブル | `table_get`・`table_edit` |
| 図形 | `button_get`・`button_edit` |
| 印刷・出力 | `page_setup_get`・`page_setup_edit` |
| VBA | 一覧と参照設定の読み取りは`vba_get`、反映は`vba_sync`、取り出しは`vba_extract`、コンパイルの確認は`vba_check`、参照設定の追加と解除は`vba_edit` |
| 実行・テスト | マクロの実行は`macro_run`、実行中の状態は`session_status`、テストは`test_run`。ダイアログへの事前の応答は、`macro_run`と`test_run`の引数 |

## 各ツールの契約

### `server_info`（段階0）

- 用途: サーバーの名前、版、ビルドの識別子を返します。AIエージェントと人が、どのビルドが動いているかを確かめるために使います。
- 引数: ありません。引数を渡すと`INVALID_REQUEST`で拒否します。
- 結果: `{"name":"xlwright","version":"0.1.0","build_id":"<commitの短いハッシュ>[-dirty]-<ビルドの時刻（UTC、yyyyMMddTHHmmssZ）>"}`。`-dirty`は、作業ツリーに差分がある時だけ付きます。
- 注釈: `ReadOnly`は真、`Destructive`は偽、`Idempotent`は真、`OpenWorld`は偽です。
- Excelに触れず、`--no-excel`でも同じに答えます。
