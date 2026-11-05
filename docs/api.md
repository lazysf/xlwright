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

- 段階4で、`macro_run`の引数と`test_run`のテストの中に、ダイアログへの事前の応答を加えます（[段階4の契約](#ダイアログへの事前の応答)）。
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
| 実行・テスト | マクロの実行は`macro_run`、実行中の状態は`session_status`、テストは`test_run`。ダイアログへの事前の応答は、`macro_run`の引数と、`test_run`のテストの中のVBA |

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
- アプリの排他、所有記録、開始の順序の4の検査、次の開始での片付けは、段階2で作りました。

### `session_start`（段階1）

- 用途: アプリのフォルダを渡して、専用のExcelを起動し、主ブックを開きます。
- 引数: `app`（文字列、必須）。アプリのフォルダです。リポジトリ（サーバーの`--project`）からの相対パスか、絶対パスです。
- 結果: `{"session":{"state":"ready","busy":false},"app":"apps/sales","workbook":"Sales.xlsm","excel":{"pid":12345,"started_at_utc":"…"}}`。`app`はリポジトリからの相対パスで、区切りは`/`です。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `PATH_NOT_ALLOWED` | リポジトリの外、UNC、デバイスパス、ネットワークドライブ（[パスの制限](architecture.md#パスの制限)） |
| `NOT_FOUND` | `APP_NOT_FOUND` | フォルダが無い |
| `INVALID_REQUEST` | `NO_WORKBOOK` | 直下に主ブックが無い。段階3まで。段階4からは、主ブックの無いセッションを始める（[段階4の契約](#主ブックの無いセッション)） |
| `INVALID_REQUEST` | `MULTIPLE_WORKBOOKS` | 直下にブックが2つ以上ある。`data.workbooks`に名前の一覧 |
| `INVALID_REQUEST` | `NOT_IMPLEMENTED` | `--no-excel`で動いている。上の検査の後に返す |
| `CONFLICT` | `SESSION_ACTIVE` | このサーバーで既にセッションを始めている |
| `CONFLICT` | `SERVER_UNUSABLE` | 上の共通の扱いのとおり |
| `CONFLICT` | `WORKBOOK_READ_ONLY` | 開いた主ブックの`ReadOnly`が真。専用のExcelは終えて、`closed`に戻る |
| `EXCEL_ERROR` | `EXCEL_NOT_STARTED` | Excelを起動できなかった（`CreateInstance`の失敗）。`closed`のまま（人が承認、2026-11-04） |
| `EXCEL_ERROR` | `OPEN_FAILED` | Excelが主ブックを開けなかった。専用のExcelは終えて、`closed`に戻る |
| `STATE_UNKNOWN` | `OWNERSHIP_NOT_CONFIRMED` | 起動したExcelの所有を確定できなかった。そのExcelには触らず、`excel_ended`は偽 |
| `STATE_UNKNOWN` | `TIMEOUT` | 開始の期限（120秒）が切れた |

- 注釈: `ReadOnly`は偽、`Destructive`は偽、`Idempotent`は偽、`OpenWorld`は偽です。
- `.xlsx`の主ブックでも始められます。`vba_sync`と`macro_run`は拒否します。

### `session_start`に段階2で加えるもの（人が承認、2026-11-04）

[開始の順序](architecture.md#開始の順序)の2〜4と、起動前の印、[片付け](architecture.md#残った専用のexcelの片付け)を加えます。引数は変えません。

- 結果: 前のセッションの専用のExcelを片付けた時だけ、`cleanup`を加えます。`{"cleanup":{"excel":{"pid":12345,"started_at_utc":"…"},"ended_by":"quit"}}`。`ended_by`は、Quitで終わった時に`quit`、同一性を照合して終了させた時に`terminate`です。AIエージェントが、人の画面から消えたExcelを報告できるようにするためです。
- 片付けた後に開始が失敗した時も、失敗の`data`に`cleanup`を同じ形で入れます（人が承認、2026-11-04）。片付けの後は所有記録が完了になり、次の開始では報告されないためです。
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
- `range`は、1つのセルか1つの矩形だけです。列（`A:A`）・行・複数の範囲・名前・シートの修飾は、`INVALID_RANGE`で拒否します（`range_write`も同じ。人が承認、2026-11-04）。
- 10,000セルを超える範囲は、Excelに触れる前に`TOO_LARGE`で拒否します。1つのセルが応答で2文字以上になり、20,000文字に収まらないためです（人が承認、2026-11-04）。
- 結果: `{"session":{…},"sheet":"Data","range":"A1:C2","values":[[1,"a",true],[null,{"error":"#DIV/0!"},2.5]]}`。`content`が`"formulas"`の時は`values`の代わりに`formulas`（文字列の2次元の配列、英語の関数名のA1形式）を返します。
- 値: 数値、文字列、真偽値、空は`null`、セルのエラーは`{"error":"#N/A"}`の形です。日付は、シリアル値の数値で返します（`Value2`）。名前の分からないエラーは、`{"error":"#ERROR(<番号>)"}`で返します。名前を付けるのは、Microsoftが`XlCVError`に挙げる8つと、実機で見た`#CALC!`です（人が承認、2026-11-04）。
- 失敗: `NOT_FOUND`・`SHEET_NOT_FOUND`、`INVALID_REQUEST`・`INVALID_RANGE`、`INVALID_REQUEST`・`TOO_LARGE`（応答が20,000文字を超える。範囲の絞り方を`message`に書く）、`STATE_UNKNOWN`・`TIMEOUT`（30秒）。
- 注釈: `ReadOnly`は真、`Destructive`は偽、`Idempotent`は真、`OpenWorld`は偽です。

### `range_write`（段階1）

- 用途: 1つのシートの1つの範囲に、値か数式を書くか、中身を消します。保存はしません。
- 引数: `sheet`（文字列、必須）、`range`（A1形式、必須）。加えて、次のどれか1つが必須です。
  - `values`: 2次元の配列。要素は数値、文字列、真偽値、`null`（空にする）。
  - `formulas`: 文字列の2次元の配列。英語の関数名のA1形式です（`Range.Formula`）。`=SEQUENCE(2)`のような動的配列の数式はspillしません。`Formula2`は使いません（人が承認、2026-11-04、[記録](notes/range-values.md)）。
  - `clear`: `true`。値と数式を消し、書式は残します（`ClearContents`）。
- 配列の行と列の数は、範囲の大きさと一致させます。
- `values`の文字列は、人がセルに入力した時と同じにExcelが解釈します。`"1/2"`は日付に、`"001"`は数値になりえます。`=`で始まる文字列は、`INVALID_REQUEST`・`VALUE_LOOKS_LIKE_FORMULA`で拒否し、`formulas`を使うよう`message`に書きます。
- 書く間は`EnableEvents`を偽にします。`Worksheet_Change`などのイベントは動きません。
- 結果: `{"session":{…},"sheet":"Data","range":"A1:C2","cells":6}`。
- 失敗: `NOT_FOUND`・`SHEET_NOT_FOUND`、`INVALID_REQUEST`・`INVALID_RANGE`、`INVALID_REQUEST`・`SHAPE_MISMATCH`、`INVALID_REQUEST`・`VALUE_LOOKS_LIKE_FORMULA`、`EXCEL_ERROR`・`WRITE_FAILED`（シートの保護など。書けた範囲は分からないため、`message`で読み直しを求める）、`STATE_UNKNOWN`・`TIMEOUT`（30秒）。
- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は真、`OpenWorld`は偽です。

### `vba_sync`（段階1）

- 用途: `src/modules/*.bas`と`src/classes/*.cls`を主ブックへ全置換で反映します（[VBAの反映](architecture.md#vbaの反映)）。ソースに無い標準・クラスモジュールは、ブックから消えます。段階3で、`src/workbook/<CodeName>.cls`をシートとブックのモジュールへ書くよう広げました（[段階3の契約](#vba_syncをシートとブックのモジュールに広げる)）。保存はしません。
- 引数: ありません。
- 反映の前に、`src/modules/XlWright.bas`を[xlwrightのVBAのモジュール](architecture.md#xlwrightのvbaのモジュール)のとおりに書きます。
- 結果: `{"session":{…},"modules":[{"name":"Calc","type":"standard"},{"name":"Counter","type":"class"}],"removed":["Old"]}`。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `NOT_MACRO_ENABLED` | 主ブックが`.xlsx` |
| `INVALID_REQUEST` | `SOURCE_INVALID` | ブックに触れる前の検査で拒否した。`data.file`・`data.line`・`data.reason`。理由は、ANSIに変換できない文字、`VB_Name`とファイル名の食い違い、拡張子とフォルダの食い違い、入れ子のフォルダ、`XlWright`という名前のほかのモジュール、`XlWrightCall`という名前のモジュール（人が承認、2026-11-04） |
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
| `EXCEL_ERROR` | `MACRO_ERROR` | マクロの実行時エラーを、xlwrightのVBAのモジュールが捕まえた。`data.number`・`data.description`・`data.source`。`ready`のまま。`End`文で止まった時のように`Application.Run`自体がCOMの例外になった時は、`data.number`にHRESULTを、`data.source`に空を入れる（人が承認、2026-11-04、[記録](notes/macro-run.md#end文)） |
| `STATE_UNKNOWN` | `DIALOG` | マクロの実行中にダイアログで止まった。専用のExcelを終えて`closed`。`data.dialog`に題名と本文。前の保存からの変更は失われる |
| `STATE_UNKNOWN` | `TIMEOUT` | 実行の上限を過ぎた。専用のExcelを終えて`closed`。前の保存からの変更は失われる |

- ダイアログの文字は応答にだけ返し、診断の記録には書きません。
- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は偽、`OpenWorld`は真です。マクロはアプリのフォルダの外にも触れうるためです。

### 段階3の契約（人が承認、2026-11-04）

段階3の確認（[記録](notes/test-copy.md)）を受けた契約です。段階1のツールに共通の扱いは、段階3のツールにも当てはめます。

#### `vba_sync`をシートとブックのモジュールに広げる

- `src/workbook/<CodeName>.cls`のコードを、ブックの同じCodeNameのシートとブックのモジュールへ書きます（[VBAの反映](architecture.md#vbaの反映)）。ヘッダーと`Attribute`の行を除き、モジュールのコードをすべて置き換えます。
- CodeNameは、`VBComponents`の名前で照合します。`Worksheet.CodeName`は、`VBProject`に触れる前は空のことがあるためです。
- `src/workbook/`にファイルの無いシートとブックのモジュールには触れません。コードを消すには、空のファイルを置きます。
- ファイルと同じ検査（ANSIに変換できない文字など）を通します。
- 結果の`modules`に、`{"name":"Sheet1","type":"document"}`の形で加えます。
- 失敗: 段階2までの`WORKBOOK_MODULES_NOT_SUPPORTED`（`src/workbook/`にファイルがある）を除き、次を加えます。

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `NOT_FOUND` | `WORKBOOK_MODULE_NOT_FOUND` | `src/workbook/`のファイル名のCodeNameが、ブックのシートとブックのモジュールに無い。ブックを変える前に止まる。`data.file` |

#### `vba_get`

- 用途: 主ブックのVBAのモジュールの一覧を返します。反映はせず、今のブックの中身を返します。
- 引数: ありません。
- 結果: `{"session":{…},"modules":[{"name":"Calc","type":"standard","lines":12},{"name":"Counter","type":"class","lines":30},{"name":"ThisWorkbook","type":"document","lines":0},{"name":"Sheet1","type":"document","lines":5,"sheet":"データ"},{"name":"Form1","type":"form","lines":20}]}`。`type`は`standard`・`class`・`document`・`form`です。シートのモジュールには、シートの名前を`sheet`に返します。CodeNameとシートの名前の対応を、AIエージェントが`src/workbook/`のファイル名に使うためです。
- 失敗: `vba_sync`の`NOT_MACRO_ENABLED`・`VBA_ACCESS_NOT_TRUSTED`・`VBA_PROJECT_LOCKED`と、`STATE_UNKNOWN`・`TIMEOUT`（30秒）。
- 注釈: `ReadOnly`は真、`Destructive`は偽、`Idempotent`は真、`OpenWorld`は偽です。

#### `vba_check`

- 用途: `vba_sync`と同じ反映をしてから、主ブックのVBAのプロジェクトをコンパイルします。保存はしません。
- 引数: ありません。
- 手順: 主ブックのプロジェクトを`ActiveVBProject`にし、VBEの「VBAProject のコンパイル」（`FindControl(1, 578)`）を`Execute`します。その後、5秒の間、[ダイアログ](architecture.md#失敗の返し方と記録)を見張ります。ダイアログが出たら、文字と、`ActiveCodePane`の選択の位置を読み、ダイアログを`WM_CLOSE`で閉じ、VBEのウィンドウを隠します。
- 説明に書くこと: コンパイルのエラーでは、非表示のExcelでも、ダイアログとVBEのウィンドウが人の画面に短い間表示されます。
- 結果: コンパイルのエラーは、ツールの失敗でなく結果として返します。確認そのものは行えたためです。
  - 通った時: `{"session":{…},"compiled":true,"sync":{"modules":[…],"removed":[…]}}`。
  - エラーの時: `{"session":{…},"compiled":false,"error":{"message":"コンパイル エラー: 変数が定義されていません。","module":"Calc","line":4},"sync":{…}}`。`module`と`line`は、読めた時だけ付きます（実機では読めないことがあった）。
- 失敗: `vba_sync`の失敗（反映の前に止まった時はコンパイルしない）に加えて、次です。

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `EXCEL_ERROR` | `COMPILE_NOT_CONFIRMED` | ダイアログが出ず、コンパイルの済んだ状態（コントロールの`Enabled`が偽）にもならなかった。`ready`のまま |
| `STATE_UNKNOWN` | `DIALOG` | 閉じたダイアログが1秒の後も残る。Quitで終わらないことがあったため、専用のExcelを終えて`closed`。`data.dialog`に題名と本文 |
| `STATE_UNKNOWN` | `TIMEOUT` | 30秒の期限が切れた |

- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は真、`OpenWorld`は偽です。

#### `test_run`

- 用途: [テストの流れ](architecture.md#テスト)のとおり、反映して主ブックを保存し、テストの写しでテストを実行し、結果を返します。
- 説明に書くこと: 呼ぶ前の変更を含めて主ブックを保存すること。テストの中の変更は主ブックに残らないこと。テストがアプリのフォルダの外や、主ブックのフォルダを絶対パスで書き換えることは防がないこと。
- 引数:
  - `tests`（文字列の配列、任意）: 実行するテストの`<モジュール>.<プロシージャ>`。省くと、すべてのテストです。
  - `timeout_seconds`（整数、任意）: 1〜1200、既定は300です。保存から開き直しまでの、呼出しの全体で測ります。
- テスト: `tests/*.bas`の、引数の無い`Sub`で、`Private`でなく、名前が`Test`で始まるものです。xlwrightはソースのファイルから見つけ、ファイル名の順、ファイルの中の順に1つずつ実行します。`tests/*.cls`は、テストの補助のクラスとして取り込みます。
- 写しには、`tests/`のモジュールと、アサーションのモジュール`XlWrightAssert`を取り込みます。取り込むファイルは、`vba_sync`と同じ検査を通します。`src/`と同じ名前のモジュール、`XlWright`・`XlWrightCall`・`XlWrightAssert`という名前のモジュールは、`SOURCE_INVALID`で拒否します。
- アサーション: `XlWrightAssert.AreEqual expected, actual[, message]`、`XlWrightAssert.IsTrue condition[, message]`、`XlWrightAssert.Fail message`。失敗すると、`Err.Source`が`XlWrightAssert`のエラーを上げ、そのテストはそこで止まります。値の比較はVBAの`=`で行い、オブジェクトと配列は比べずに失敗にします。テストが`On Error`でアサーションの失敗を握りつぶすと、`passed`になります。
- テストは、`macro_run`と同じく、名前で直接呼ぶ関数を実行ごとに作って呼びます。実行の間は`EnableEvents`を変えません。
- 結果: `{"session":{…},"summary":{"passed":2,"failed":1,"errors":1},"tests":[{"name":"Tests.TestAdd","result":"passed","elapsed_ms":3},{"name":"Tests.TestSplit","result":"failed","message":"AreEqual: expected 3, actual 2. 列の数"},{"name":"Tests.TestLoad","result":"error","number":53,"description":"ファイルが見つかりません。","source":"VBAProject"}],"elapsed_ms":1520,"sync":{…}}`。テストが失敗しても、ツールは成功を返します。各テストには`elapsed_ms`も付きます。結果が大きい時（18,000文字を超える時）は、全体を`<アプリ>/.xlwright/results/<id>.json`へ書いて`result_file`にそのパスを返し、`tests`には`passed`でないテストを収まるだけ返します。
- 後始末: 写しのフォルダ（`.xlwright/test/<run-id>/`）の中のブックは、保存せずに閉じます。ほかのブックは閉じません（`session_end`の`FOREIGN_WORKBOOK`で知らせます）。写しのフォルダを消せない時は、結果の`leftover_copy`にそのパス（リポジトリからの相対パスで、区切りは`/`）を返します。
- 写しへのテストの取り込みに失敗した時は、主ブックを開き直してから、`vba_sync`と同じ`SYNC_FAILED`を返します。
- 失敗: `vba_sync`の失敗（反映の前に止まった時は実行しない）に加えて、次です。主ブックを閉じた後に止まった時は、主ブックを開き直してから返します。開き直せない時は、下の`WORKBOOK_NOT_REOPENED`です。止まるまでのテストの結果は、`data.tests`に入れます。

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `INVALID_TEST_NAME` | `tests`の名前の形が違う |
| `NOT_FOUND` | `TEST_NOT_FOUND` | テストが1つも無い、または`tests`の名前のテストが無い。Excelに触れる前に止まる |
| `EXCEL_ERROR` | `SAVE_FAILED` | 主ブックを保存できなかった。テストを実行せず、`ready`のまま |
| `INTERNAL_ERROR` | `COPY_FAILED` | アプリのフォルダを写せなかった |
| `EXCEL_ERROR` | `OPEN_FAILED` | 写しを開けなかった |
| `CONFLICT` | `WORKBOOK_NOT_REOPENED` | 主ブックを開き直せなかった、または`ReadOnly`で開いた。専用のExcelを[終了](architecture.md#終了)の2〜5で終え、`closed`。主ブックは保存してある |
| `STATE_UNKNOWN` | `DIALOG` | テストの実行中にダイアログで止まった（コンパイルのエラーを含む）。専用のExcelを終えて`closed`。`data.dialog`と、止まったテストの名前を`data.test`に入れる。主ブックは保存してある |
| `STATE_UNKNOWN` | `TIMEOUT` | 実行の上限を過ぎた。専用のExcelを終えて`closed`。`data.test`は上と同じ |

- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は偽、`OpenWorld`は真です。

#### `session_status`の実行中のテスト

- `test_run`の実行中は、`operation`に`test`を加えます。`{"tool":"test_run","elapsed_ms":4210,"test":{"name":"Tests.TestLoad","index":3,"count":5}}`。`index`は1から数えます。写しを開く前と開き直す間は、`test`を付けません。

### 段階4の契約（人が承認、2026-11-05）

段階4の確認（[記録](notes/vba-extract.md)）と、[要件の範囲](requirements.md#操作)のブック・VBA・実行とテストの行を受けた契約です。`test_run`の応答をテストの中で決めることと、`test_run`の前のコンパイルも、人が選びました。段階1のツールに共通の扱いは、段階4のツールにも当てはめます。

#### 主ブックの無いセッション

- `session_start`は、直下にブックが無い時に`NO_WORKBOOK`で拒否せず、主ブックの無いセッションを始めます（[アプリのフォルダ](architecture.md#アプリのフォルダ)）。専用のExcelを起動し、ブックは開きません。結果の`workbook`は`null`です。
- このセッションでは、`workbook_create`・`session_status`・`session_end`・`server_info`だけを受け付けます。ほかのツールは`CONFLICT`・`detail: NO_WORKBOOK`で拒否し、`workbook_create`を呼ぶよう`message`に書きます。
- `session_end`の`save`は、主ブックが無ければ無視し、結果の`saved`は偽です。
- 所有記録の`workbook`には、開始の時の主ブックの名前を書きます。主ブックの無いセッションでは空の文字列です。記録の形式は変えません。

#### `workbook_create`

- 用途: 主ブックの無いセッションで、新しいブックを作り、アプリのフォルダの直下に保存します。作ったブックが主ブックになり、ほかのツールを使えるようになります。
- 引数: `name`（文字列、必須）。ファイルの名前で、拡張子は`.xlsm`か`.xlsx`です。フォルダの区切り、`~$`で始まる名前、Windowsでファイル名に使えない文字は拒否します。
- 手順: `Workbooks.Add`で作り、`SaveAs`で拡張子に合う形式（`.xlsm`は52、`.xlsx`は51）で保存します。
- 結果: `{"session":{…},"workbook":"Sales.xlsm"}`。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `INVALID_WORKBOOK_NAME` | `name`の形が違う |
| `CONFLICT` | `WORKBOOK_EXISTS` | セッションに主ブックがある |
| `CONFLICT` | `FILE_EXISTS` | 直下に同じ名前のファイルがある |
| `EXCEL_ERROR` | `SAVE_FAILED` | 保存に失敗した。作ったブックは保存せずに閉じ、主ブックの無いセッションのまま |
| `STATE_UNKNOWN` | `TIMEOUT` | 保存の期限（120秒）が切れた |

- 注釈: `ReadOnly`は偽、`Destructive`は偽、`Idempotent`は偽、`OpenWorld`は偽です。

#### `workbook_get`と`workbook_edit`

- `workbook_get`の用途: 主ブックの名前、形式、保存していない変更の有無を返します。結果: `{"session":{…},"workbook":"Sales.xlsx","format":"xlsx","saved":true}`。`format`は`xlsm`か`xlsx`です。失敗は`STATE_UNKNOWN`・`TIMEOUT`（30秒）だけです。注釈: `ReadOnly`は真、`Destructive`は偽、`Idempotent`は真、`OpenWorld`は偽です。
- `workbook_edit`の引数: `action`（文字列、必須）。段階4では`"convert_to_xlsm"`だけです。段階5で、ブックの保護を加えます。
- `convert_to_xlsm`: `.xlsx`の主ブックを、同じ名前の`.xlsm`として直下に`SaveAs`（52）で保存します。以後、セッションの主ブックは`.xlsm`です。元の`.xlsx`は、前の保存のまま残ります。AIエージェントがGitの操作で消すまで、次の`session_start`は`MULTIPLE_WORKBOOKS`で止まります（[要件](requirements.md#操作)）。このことを`message`ではなく、ツールの説明に書きます。
- 結果: `{"session":{…},"workbook":"Sales.xlsm","previous":"Sales.xlsx"}`。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `ALREADY_MACRO_ENABLED` | 主ブックが既に`.xlsm` |
| `CONFLICT` | `FILE_EXISTS` | 直下に同じ名前の`.xlsm`がある |
| `EXCEL_ERROR` | `SAVE_FAILED` | 保存に失敗した。主ブックは`.xlsx`のまま |
| `STATE_UNKNOWN` | `TIMEOUT` | 保存の期限（120秒）が切れた |

- 注釈: `ReadOnly`は偽、`Destructive`は偽、`Idempotent`は偽、`OpenWorld`は偽です。

#### `vba_extract`

- 用途: アプリのフォルダの中の既存のブックから、VBAのソースを`src/`へ取り出します（[VBAの反映](architecture.md#vbaの反映)）。ブックの中にしか無いモジュールは次の反映で消えるため、既存のブックでは`vba_sync`の前に呼びます。
- 引数: `workbook`（文字列、任意）。アプリのフォルダからの相対パスです。省くと主ブックです。アプリのフォルダの外は`PATH_NOT_ALLOWED`で拒否します。
- 手順:
  1. ブックが`.xlsm`であることと、パッケージに`xl/macrosheets/`が無いことを確かめます。
  2. ファイルを`.xlwright/extract/<id>/`へ、`<id>.xlsm`の名前で写します。主ブックと同じ名前のブックは同時に開けないためです（[記録](notes/test-copy.md#テストの写しの流れ)）。
  3. 写しを`AutomationSecurity=ForceDisable`・`EnableEvents=False`で開きます。開いた後、`AutomationSecurity`を元に戻します。戻しても写しのマクロは無効のままでした（[記録](notes/vba-extract.md)）。
  4. 標準・クラス・シートとブックのモジュールを`Export`し、システムのANSIコードページからUTF-8へ変えて、`src/modules/<名前>.bas`・`src/classes/<名前>.cls`・`src/workbook/<CodeName>.cls`へ書きます。UserFormと`XlWright`という名前のモジュールは書きません。
  5. 写しを保存せずに閉じ、`<id>/`を消します。
- 開いているブック（主ブック）は変えません。保存もしません。
- 結果: `{"session":{…},"workbook":"Sales.xlsm","files":[{"file":"src/modules/Calc.bas","type":"standard"},{"file":"src/workbook/Sheet1.cls","type":"document"}],"skipped":[{"name":"Form1","type":"form"}]}`。`file`はアプリのフォルダからの相対パスで、区切りは`/`です。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `PATH_NOT_ALLOWED` | `workbook`がアプリのフォルダの外 |
| `NOT_FOUND` | `WORKBOOK_NOT_FOUND` | `workbook`のファイルが無い |
| `INVALID_REQUEST` | `NOT_MACRO_ENABLED` | `.xlsm`でない |
| `INVALID_REQUEST` | `EXCEL4_MACROS` | パッケージに`xl/macrosheets/`がある。開かずに止まる |
| `CONFLICT` | `SOURCE_EXISTS` | `src/`に、`.`で始まるものと`XlWright.bas`を除くファイルがある。上書きしないため。`data.files`に一覧 |
| `CONFLICT` | `VBA_ACCESS_NOT_TRUSTED` | `vba_sync`と同じ |
| `CONFLICT` | `VBA_PROJECT_LOCKED` | `vba_sync`と同じ |
| `EXCEL_ERROR` | `OPEN_FAILED` | 写しを開けなかった |
| `EXCEL_ERROR` | `EXTRACT_FAILED` | 途中で失敗した。`data.done`に書いたファイル。書いたファイルは消さない |
| `STATE_UNKNOWN` | `TIMEOUT` | 30秒の期限が切れた。開く操作は保存の期限（120秒）で測る |

- 注釈: `ReadOnly`は偽、`Destructive`は偽、`Idempotent`は偽、`OpenWorld`は偽です。

#### `vba_get`の参照設定

- 結果に`references`を加えます。`[{"name":"Scripting","guid":"{420B2830-E718-11CF-893D-00A0C9054228}","major":1,"minor":0,"builtin":false,"broken":false}]`。読めた参照設定の順です（[記録](notes/vba-extract.md#forcedisableでvbaを取り出せるか)）。`broken`が真で`name`を読めない時は、`name`を省きます。

#### `vba_edit`

- 用途: 主ブックの参照設定を追加と解除します。保存はしません。
- 引数: `action`（`"add_reference"`か`"remove_reference"`、必須）、`guid`（文字列、必須）。`add_reference`では`major`と`minor`（整数、必須）も渡します。`vba_get`の`references`の値をそのまま使えます。
- 手順: 追加は`References.AddFromGuid`、解除は`References.Remove`です。ファイルのパスからの追加（`AddFromFile`）は作りません。
- 結果: `{"session":{…},"changed":true,"references":[…]}`。同じ`guid`が既にある時の追加は、変えずに`changed`を偽で返します。
- 失敗:

| `code` | `detail` | 場合 |
| --- | --- | --- |
| `INVALID_REQUEST` | `NOT_MACRO_ENABLED` | 主ブックが`.xlsx` |
| `NOT_FOUND` | `REFERENCE_NOT_FOUND` | 解除する`guid`の参照設定が無い |
| `INVALID_REQUEST` | `BUILTIN_REFERENCE` | `builtin`が真の参照設定（`VBA`・`Excel`）の解除 |
| `CONFLICT` | `VBA_ACCESS_NOT_TRUSTED` | `vba_sync`と同じ |
| `CONFLICT` | `VBA_PROJECT_LOCKED` | `vba_sync`と同じ |
| `EXCEL_ERROR` | `REFERENCE_FAILED` | Excelが追加か解除を断った（登録されていないライブラリなど） |
| `STATE_UNKNOWN` | `TIMEOUT` | 30秒の期限が切れた |

- 注釈: `ReadOnly`は偽、`Destructive`は真、`Idempotent`は真、`OpenWorld`は偽です。

#### ダイアログへの事前の応答

[要件の利用の前提](requirements.md#利用の前提)のとおり、マクロが`XlWright.bas`の関数でダイアログを出す時だけ働きます。

- 関数: どれも、xlwrightの実行の外では標準のダイアログを出します。
  - `XlWright.Message(prompt[, buttons][, title])`: `VBA.MsgBox`の代わり。押したボタンの値を返します。
  - `XlWright.Input(prompt[, title][, default])`: 入力。文字列を、キャンセルでは`False`を返します（`Application.InputBox`の`Type:=2`）。
  - `XlWright.OpenFile([filter][, title])`・`XlWright.SaveFile([initial][, filter][, title])`: ファイルの選択と保存先の指定。完全なパスを、キャンセルでは`False`を返します（`Application.GetOpenFilename`・`GetSaveAsFilename`）。
  - `XlWright.Folder([title])`: フォルダーの選択。完全なパスを、キャンセルでは`False`を返します（`FileDialog`のフォルダーの選択）。
  - VBAの`MsgBox`と同じ名前にしないのは、修飾の無い`MsgBox`の呼出しをすべて置き換えてしまうためです。
- `macro_run`の引数に`dialogs`（配列、任意、30個まで）を加えます。要素は順に使う応答です。
  - `{"kind":"message","answer":"yes"}`。`answer`は`ok`・`cancel`・`yes`・`no`・`abort`・`retry`・`ignore`です。
  - `{"kind":"input","value":"東京"}`、`{"kind":"open_file","path":"data/in.csv"}`、`{"kind":"save_file","path":"data/out.csv"}`、`{"kind":"folder","path":"data"}`。`path`はアプリのフォルダからの相対パスで、外を指すと`PATH_NOT_ALLOWED`で拒否します。xlwrightが完全なパスに変えて渡します。
  - どの`kind`でも、`answer`・`value`・`path`の代わりに`"cancel":true`を渡せます。
- xlwrightの実行の中では、`dialogs`を省いても、関数はダイアログを出しません。応答が尽きた時、次の応答の`kind`が違う時、`message`の`answer`が`buttons`に無い時は、`Err.Source`が`XlWright`のエラーを上げます。マクロは`MACRO_ERROR`で止まり、専用のExcelは終えません。
- 結果に`dialogs`を加えます。`{"answered":2,"unused":1}`。使われなかった応答があっても、失敗にしません。
- `test_run`では、テストの中で応答を決めます。`XlWright.Answer kind, value`をテストが呼び、そのテストの中の関数が順に使います。応答はテストごとに空にします。`test_run`の引数には加えません。テストごとに応答が違い、テストの書き手が知っているためです。

#### `test_run`の前のコンパイル

- 写しへテストを取り込んだ後、テストを実行する前に、`vba_check`と同じコンパイルを行います。エラーの時はテストを実行せず、写しを閉じて主ブックを開き直し、`EXCEL_ERROR`・`detail: COMPILE_ERROR`と`data.error`（`vba_check`の`error`と同じ形）を返します。`ready`のままです。
- まだ無い関数を呼ぶテストで、専用のExcelを終えずに返すためです（[記録](notes/stage3-scenarios.md#分かったこと)）。

## 表示要求のコマンド（段階2、人が承認、2026-11-04）

[人への引継ぎ](architecture.md#人への引継ぎ)で、残った専用のExcelを人へ見せ、人が確かめて閉じた後に終えるコマンドです。MCPのツールではなく、人がPowerShell 5.1から実行します。形は人が決めました（2026-11-04）: サーバーの実行ファイルのサブコマンドにし、排他・所有記録の読み・プロセスの照合はCoreの部品を使います。人が閉じた後もプロセスが残る時は、コマンドが照合して終了させます。

- 開発版の入口: `powershell -NoProfile -File plugins\xlwright-dev\scripts\show-excel.ps1 -App <アプリのフォルダ>`。スクリプトは、`.xlwright/builds/latest.json`のビルドのサーバーを、下の引数で起動するだけです。利用する版の入口は段階4で決めます。
- 引数: `show --project <リポジトリ> --app <アプリのフォルダ>`。`--app`は、`session_start`の`app`と同じ扱いです（[パスの制限](architecture.md#パスの制限)）。外れた時は、使い方を標準エラーに書いて終了コード2で終わります。
- 出力: 人への説明を、英語のASCIIで標準出力に書きます。PowerShell 5.1の既定のコンソールの文字コードで崩れないためです。
- 終了コード: 0は残った専用のExcelが無い（初めから無いか、終えた）。1は止まった（使用中、記録の不正、確かめられないExcel、終了を確かめられない）。2は引数の不正です。

手順:

1. パスの制限を確かめます。
2. 所有記録のファイルが無ければ、作らずに、記録が無いと書いて0で終わります。
3. アプリの排他を取ります。取れなければ、別の会話が使用中と書いて1で終わります。
4. 所有記録を、開始と同じく厳密に読みます。空・不正・読めない時は、[人への引継ぎ](architecture.md#人への引継ぎ)の手順を書いて1で終わります。
   - `completed`: 0で終わります。
   - `starting`: 印の時刻と、その後にCOMから起動されたEXCELのPIDと生成時刻を書き、見せずに1で終わります。記録したPIDが無いためです。
   - `owned`: 記録の同一性をOSのプロセスと照合します。プロセスが無ければ0で終わります。照会できない時と、コマンドラインに`/automation`が無いか読めない時は、触らずに1で終わります。
5. 一致したプロセスの、`EXCEL7`を持つ`XLMAIN`に、`ShowWindow`と`SetForegroundWindow`を呼びます。COMは使いません。`EXCEL7`が無い時は、見せるブックのウィンドウが無いと書きます。
6. 人が`done`と入力するまで、排他を持ち続けます。標準入力が閉じた時は、終了させずに1で終わります。
7. `done`の後、プロセスが終わっていれば0で終わります。残っていて`EXCEL7`が無ければ、照合に使ったハンドルで終了させ、終了を確かめて0で終わります。確かめられなければ1です。`EXCEL7`が残っていれば、終了させず、まだブックが開いていると書いて、6へ戻ります。
8. 所有記録は書き換えません。次の開始が、プロセスの不在を確かめて完了にします。排他を外して終わります。

- 終了させる根拠は、片付けで`EXCEL7`が無い時と同じです。排他を取れていること、同一性と`/automation`を確かめていること、ブックのウィンドウが無いことです（[片付け](architecture.md#残った専用のexcelの片付け)）。
