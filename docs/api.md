# API

本書は、xlwrightのMCPのツールの名前と共通の規約を定めます。
ツールの名前・説明・引数・応答は、公開インターフェースです。モデルはツールの定義を、指示として読みます。
各ツールの引数と応答は、そのツールを作る段階で、その段階の実機の確認の直後に本書へ書きます。実装の前に人の承認を得ます。
方式は[設計](architecture.md)にあります。

## 共通の規約

### 名前と分け方

ツールは、読むツールと変えるツールに分けます。
読むツールに、読み取りだけの注釈を正しく付けられ、モデルが名前を予測しやすくなるためです。範囲の表の操作（およそ70）を操作ごとのツールにすると数が多すぎ、対象ごとに1つのツールにすると読み取りと削除が混ざります。

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
- 失敗は、`isError: true`の`CallToolResult`で返します。`structuredContent`は付けません。MCPの仕様では、ツールの実行の失敗をプロトコルのエラーではなく結果として返し、モデルが読んで次の行動を決められるようにするためです（[Tools](https://modelcontextprotocol.io/specification/draft/server/tools)）。
- 失敗の本文は、次の形の1行のJSONです。

```json
{"error":{"code":"CONFLICT","detail":"SESSION_IN_USE","message":"…","data":{}},"session":{"state":"closed","busy":false}}
```

- `message`には、次に取る行動を書きます。
- `data`は、ツールごとの補足です。`STATE_UNKNOWN`以外では省略できます。部分的に変更した時の処理済みの部分、止まって知らせる時の引継ぎの情報（記録したPIDと生成時刻、所有記録のパス、人が取る手順）、診断の記録のパスを入れます。中身は各ツールの契約で定めます。
- `code`は次の6つです。`detail`の値は、各ツールの契約で定めます。契約が`detail`の値を定めない失敗では、`detail`を省きます。

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
- 応答の本文は20,000文字までです。本文と結果のファイルは、ASCII以外の文字をescapeせずに書きます。日本語の1文字は1文字に数えます。
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
- 結果: `{"name":"xlwright","version":"0.1.0","build_id":"<commitの短いハッシュ>[-dirty]-<ビルドの時刻（UTC、yyyyMMddTHHmmssZ）>"}`。`-dirty`は、作業ツリーに差分がある時だけ付きます。確認の入口を通らないビルドでは、`build_id`は`unknown`です。
- 注釈: `ReadOnly`は真、`Destructive`は偽、`Idempotent`は真、`OpenWorld`は偽です。
- Excelに触れず、`--no-excel`でも同じに答えます。

### 段階1のツールに共通の扱い

段階1の8つのツールの契約は、人が承認しました（2026-11-03）。

- 引数の名前と`detail`の値は、英小文字の`snake_case`と英大文字の`SNAKE_CASE`です。
- `session_start`と`session_status`のほかは、セッションが`closed`の時に`CONFLICT`・`detail: NO_SESSION`で拒否します。
- `session_status`のほかは、実行中（`busy`が真）の時に待たせず、`CONFLICT`・`detail: BUSY`で拒否します。
- 成功の結果には、どれも`session`（`{"state":…,"busy":…}`）を含めます。
- 期限は[設計の期限の表](architecture.md#期限と大きさの上限)のとおりです。期限が切れた時は`STATE_UNKNOWN`・`detail: TIMEOUT`です。
- `STATE_UNKNOWN`の`data`には、`excel_ended`と`restartable`に加えて、専用のExcelを残して止まった時は`excel`（`pid`と`started_at_utc`）と人が取る手順を入れます。
- 失敗の`data.log`に、診断の記録のパスを入れます（[失敗の返し方と記録](architecture.md#失敗の返し方と記録)）。
- 保留中のCOM呼出しが残って、このサーバーで新しいセッションを始められない時は、`session_start`を`CONFLICT`・`detail: SERVER_UNUSABLE`で拒否し、会話を閉じて新しい会話で始めるよう`message`に書きます（[状態と結果不明](architecture.md#状態と結果不明)）。
- アプリの排他、所有記録、開始の順序の4の検査は、段階2で作りました。次の開始での片付けは、まだありません。残った専用のExcelが動いている時は、`LEFTOVER_EXCEL`で止まります。

### `session_start`（段階1）

- 用途: アプリのフォルダを渡して、専用のExcelを起動し、主ブックを開きます。
- 引数: `app`（文字列、必須）。アプリのフォルダです。リポジトリ（サーバーの`--project`）からの相対パスか、絶対パスです。
- 結果: `{"session":{"state":"ready","busy":false},"app":"apps/sales","workbook":"Sales.xlsm","excel":{"pid":12345,"started_at_utc":"…"}}`。`app`はリポジトリからの相対パスで、区切りは`/`です。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `PATH_NOT_ALLOWED` | リポジトリの外、UNC、デバイスパス、ネットワークドライブ（[パスの制限](architecture.md#パスの制限)） |
| `NOT_FOUND` | `APP_NOT_FOUND` | フォルダが無い |
| `INVALID_REQUEST` | `NO_WORKBOOK` | 直下に主ブックが無い（主ブックの無いセッションは段階4） |
| `INVALID_REQUEST` | `MULTIPLE_WORKBOOKS` | 直下にブックが2つ以上ある。`data.workbooks`に名前の一覧 |
| `INVALID_REQUEST` | `NOT_IMPLEMENTED` | `--no-excel`で動いている。上の検査の後に返す |
| `CONFLICT` | `SESSION_ACTIVE` | このサーバーで既にセッションを始めている |
| `CONFLICT` | `SERVER_UNUSABLE` | 上の共通の扱いのとおり |
| `CONFLICT` | `WORKBOOK_READ_ONLY` | 開いた主ブックの`ReadOnly`が真。専用のExcelは終えて、`closed`に戻る |
| `EXCEL_ERROR` | `OPEN_FAILED` | Excelが主ブックを開けなかった。専用のExcelは終えて、`closed`に戻る |
| `STATE_UNKNOWN` | `OWNERSHIP_NOT_CONFIRMED` | 起動したExcelの所有を確定できなかった。そのExcelには触らず、`excel_ended`は偽 |
| `STATE_UNKNOWN` | `TIMEOUT` | 開始の期限（120秒）が切れた |

- 注釈: `ReadOnly`は偽、`Destructive`は偽、`Idempotent`は偽、`OpenWorld`は偽です。
- `.xlsx`の主ブックでも始められます。`vba_sync`と`macro_run`は拒否します。

### `session_start`に段階2で加えるもの（人が承認、2026-11-04）

[開始の順序](architecture.md#開始の順序)の2〜4と、起動前の印、[片付け](architecture.md#残った専用のexcelの片付け)を加えます。引数は変えません。

- 結果: 前のセッションの専用のExcelを片付けた時だけ、`cleanup`を加えます。`{"cleanup":{"excel":{"pid":12345,"started_at_utc":"…"},"ended_by":"quit"}}`。`ended_by`は、Quitで終わった時に`quit`、同一性を照合して終了させた時に`terminate`です。AIエージェントが、人の画面から消えたExcelを報告できるようにするためです。
- 止まって知らせる失敗の`data`には、`record`（所有記録のパス。リポジトリからの相対パスで、区切りは`/`）を入れます。残ったExcelがある時は、`excel`（`pid`と`started_at_utc`）も入れます。人が取る手順は、[人への引継ぎ](architecture.md#人への引継ぎ)のとおり`message`に書きます。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `CONFLICT` | `SESSION_IN_USE` | アプリの排他を取れない。別の会話がそのアプリのセッションを持っている。その会話で`session_end`を呼ぶか、その会話を閉じるよう`message`に書く |
| `CONFLICT` | `RECORD_INVALID` | 所有記録が空、形式が不正、または読めない。`data.record`を入れる |
| `CONFLICT` | `LEFTOVER_EXCEL` | 残った専用のExcelを確かめられず、触らずに止まった。`data.record`と`data.excel`を入れる |
| `CONFLICT` | `UNKNOWN_EXCEL` | 起動前の印だけが残り、その後にCOMから起動されたExcelがある。`data.record`と、`data.processes`（`pid`と`started_at_utc`の配列）を入れる |
| `CONFLICT` | `EXIT_NOT_CONFIRMED` | 残った専用のExcelを終えようとしたが、プロセスの終了を確かめられなかった。`data.record`と`data.excel`を入れる |
| `CONFLICT` | `WORKBOOK_IN_USE` | 主ブックのファイルを書込みのアクセスで開けない（共有違反）。ほかのプログラムが編集できる状態で開いている。閉じるよう`message`に書く |
| `CONFLICT` | `WORKBOOK_FILE_READ_ONLY` | 主ブックのファイルの属性が読み取り専用 |
| `STATE_UNKNOWN` | `TIMEOUT` | 片付けの照合を終える前に開始の期限が切れた時は、残ったExcelに触らず、`excel_ended`は偽、`data.record`と`data.excel`を入れる |

- どの失敗でも、このセッションの専用のExcelはまだ起動していません。`STATE_UNKNOWN`の`TIMEOUT`のほかは、`session`は`closed`のままで、`data.excel_ended`は付けません。

### `session_status`（段階1）

- 用途: セッションの状態と、実行中の操作を返します。COMを使わないため、マクロの実行中も答えます。
- 引数: ありません。
- 結果: セッションがある時は`{"session":{"state":"ready","busy":true},"app":"apps/sales","workbook":"Sales.xlsm","operation":{"tool":"macro_run","elapsed_ms":12345}}`。`operation`は実行中の時だけ付きます。セッションが無い時は`{"session":{"state":"closed","busy":false}}`です。
- 失敗: 引数を渡した時の`INVALID_REQUEST`だけです。
- 注釈: `ReadOnly`は真、`Destructive`は偽、`Idempotent`は真、`OpenWorld`は偽です。

### `session_end`（段階1）

- 用途: 主ブックを保存するか決めて閉じ、専用のExcelを終えます（[終了](architecture.md#終了)）。報告の前に呼びます。
- 引数: `save`（真偽値、必須）。偽なら、前の保存からの変更を捨てます。
- 結果: `{"session":{"state":"closed","busy":false},"saved":true}`。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `EXCEL_ERROR` | `SAVE_FAILED` | 保存に失敗した。閉じずに`ready`のまま。`save`を偽にして呼び直せる |
| `CONFLICT` | `FOREIGN_WORKBOOK` | アプリのフォルダの外のブックが開いている。専用のExcelを残して`closed`にし、`data`に保存の結果と引継ぎの情報を入れる |
| `CONFLICT` | `EXIT_NOT_CONFIRMED` | プロセスの終了を確かめられなかった。`data`は上と同じ |
| `STATE_UNKNOWN` | `TIMEOUT` | 保存の期限（120秒）が切れた。主ブックのファイルが保存の前と後のどちらかは分からない |

- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は偽、`OpenWorld`は偽です。

### `workbook_save`（段階1）

- 用途: 主ブックを上書き保存します。セッションは続きます。
- 引数: ありません。
- 結果: `{"session":{"state":"ready","busy":false},"workbook":"Sales.xlsm"}`。
- 失敗: `EXCEL_ERROR`・`SAVE_FAILED`（`ready`のまま）、`STATE_UNKNOWN`・`TIMEOUT`（120秒）。
- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は真、`OpenWorld`は偽です。

### `range_read`（段階1）

- 用途: 1つのシートの1つの範囲の値か数式を読みます。読む前に再計算し、計算を終えた値を返します（[起動時の設定](architecture.md#起動時の設定)）。
- 引数: `sheet`（文字列、必須）、`range`（A1形式、必須。例`A1:C10`）、`content`（`"values"`か`"formulas"`、既定は`"values"`）。
- 結果: `{"session":{…},"sheet":"Data","range":"A1:C2","values":[[1,"a",true],[null,{"error":"#DIV/0!"},2.5]]}`。`content`が`"formulas"`の時は`values`の代わりに`formulas`（文字列の2次元の配列、英語の関数名のA1形式）を返します。
- 値: 数値、文字列、真偽値、空は`null`、セルのエラーは`{"error":"#N/A"}`の形です。日付は、シリアル値の数値で返します（`Value2`）。
- 失敗: `NOT_FOUND`・`SHEET_NOT_FOUND`、`INVALID_REQUEST`・`INVALID_RANGE`、`INVALID_REQUEST`・`TOO_LARGE`（応答が20,000文字を超える。範囲の絞り方を`message`に書く）、`STATE_UNKNOWN`・`TIMEOUT`（30秒）。
- 注釈: `ReadOnly`は真、`Destructive`は偽、`Idempotent`は真、`OpenWorld`は偽です。

### `range_write`（段階1）

- 用途: 1つのシートの1つの範囲に、値か数式を書くか、中身を消します。保存はしません。
- 引数: `sheet`（文字列、必須）、`range`（A1形式、必須）。加えて、次のどれか1つが必須です。
  - `values`: 2次元の配列。要素は数値、文字列、真偽値、`null`（空にする）。
  - `formulas`: 文字列の2次元の配列。英語の関数名のA1形式です（`Range.Formula`）。
  - `clear`: `true`。値と数式を消し、書式は残します（`ClearContents`）。
- 配列の行と列の数は、範囲の大きさと一致させます。
- `values`の文字列は、人がセルに入力した時と同じにExcelが解釈します。`"1/2"`は日付に、`"001"`は数値になりえます。`=`で始まる文字列は、`INVALID_REQUEST`・`VALUE_LOOKS_LIKE_FORMULA`で拒否し、`formulas`を使うよう`message`に書きます。
- 書く間は`EnableEvents`を偽にします。`Worksheet_Change`などのイベントは動きません。
- 結果: `{"session":{…},"sheet":"Data","range":"A1:C2","cells":6}`。
- 失敗: `NOT_FOUND`・`SHEET_NOT_FOUND`、`INVALID_REQUEST`・`INVALID_RANGE`、`INVALID_REQUEST`・`SHAPE_MISMATCH`、`INVALID_REQUEST`・`VALUE_LOOKS_LIKE_FORMULA`、`EXCEL_ERROR`・`WRITE_FAILED`（シートの保護など。書けた範囲は分からないため、`message`で読み直しを求める）、`STATE_UNKNOWN`・`TIMEOUT`（30秒）。
- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は真、`OpenWorld`は偽です。

### `vba_sync`（段階1）

- 用途: `src/modules/*.bas`と`src/classes/*.cls`を主ブックへ全置換で反映します（[VBAの反映](architecture.md#vbaの反映)）。ソースに無い標準・クラスモジュールは、ブックから消えます。保存はしません。
- 引数: ありません。
- 反映の前に、`src/modules/XlWright.bas`を[xlwrightのVBAのモジュール](architecture.md#xlwrightのvbaのモジュール)のとおりに書きます。
- 結果: `{"session":{…},"modules":[{"name":"Calc","type":"standard"},{"name":"Counter","type":"class"}],"removed":["Old"]}`。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `NOT_MACRO_ENABLED` | 主ブックが`.xlsx` |
| `INVALID_REQUEST` | `SOURCE_INVALID` | ブックに触れる前の検査で拒否した。`data.file`・`data.line`・`data.reason`。理由は、ANSIに変換できない文字、`VB_Name`とファイル名の食い違い、拡張子とフォルダの食い違い、入れ子のフォルダ、`XlWright`という名前のほかのモジュール |
| `INVALID_REQUEST` | `WORKBOOK_MODULES_NOT_SUPPORTED` | `src/workbook/`にファイルがある（段階3） |
| `CONFLICT` | `VBA_ACCESS_NOT_TRUSTED` | 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が無効。人に設定を求める`message` |
| `CONFLICT` | `VBA_PROJECT_LOCKED` | VBAのプロジェクトが保護されている |
| `EXCEL_ERROR` | `SYNC_FAILED` | 途中で失敗した。`data.done`に処理済みのモジュール。保存していないため、次の反映がやり直す |
| `STATE_UNKNOWN` | `TIMEOUT` | 30秒の期限が切れた |

- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は真、`OpenWorld`は偽です。

### `macro_run`（段階1）

- 用途: `vba_sync`と同じ反映をしてから、主ブックのマクロを実行し、戻り値を返します（[マクロの実行](architecture.md#マクロの実行)）。マクロが終わるか上限を過ぎるまで応答しません。保存はしません。
- 引数:
  - `macro`（文字列、必須）: `<モジュール>.<プロシージャ>`の形です。ブックの修飾（`!`・`'`・`[`・パスの区切り）を含む名前は拒否します。修飾はxlwrightが付けます。
  - `args`（配列、任意）: 引数。要素は数値、文字列、真偽値です。30個までです。
  - `timeout_seconds`（整数、任意）: 1〜1200、既定は300です。
- 結果: `{"session":{…},"macro":"Calc.Twice","return":42,"elapsed_ms":15,"sync":{"modules":[…],"removed":[…]}}`。戻り値は数値・文字列・真偽値・`null`です。それ以外の型の時は、`return`を`null`にして`return_type`に型の名前を返します。戻り値が大きい時は要約し、全体を`<アプリ>/.xlwright/results/<id>.json`へ書いて`result_file`にそのパスを返します。
- 失敗: `vba_sync`の失敗（反映の前に止まった時は実行しない）に加えて、次です。

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `INVALID_MACRO_NAME` | `macro`の形が違う |
| `NOT_FOUND` | `MACRO_NOT_FOUND` | マクロが無い |
| `EXCEL_ERROR` | `MACRO_ERROR` | マクロの実行時エラーを、xlwrightのVBAのモジュールが捕まえた。`data.number`・`data.description`。`ready`のまま |
| `STATE_UNKNOWN` | `DIALOG` | マクロの実行中にダイアログで止まった。専用のExcelを終えて`closed`。`data.dialog`に題名と本文。前の保存からの変更は失われる |
| `STATE_UNKNOWN` | `TIMEOUT` | 実行の上限を過ぎた。専用のExcelを終えて`closed`。前の保存からの変更は失われる |

- ダイアログの文字は応答にだけ返し、診断の記録には書きません。
- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は偽、`OpenWorld`は真です。マクロはアプリのフォルダの外にも触れうるためです。
