# P0-core API契約

本書は実装対象の操作契約です。利用可能な実装を示すものではありません。
MCPのJSON Schemaと応答全体のスキーマは、各操作の実装前に確定します。
未確定のフィールドを、実装上の都合だけで公開仕様にしないでください。
実行制御と安全条件は [設計文書](architecture.md) に従います。
本書の全操作を初期PoCで一括実装する必要はありません。提供する範囲は [Spike-0](spike-0.md#実施順序と着手条件) で段階化します。
初期は残存Excelの自動後始末・force・保留中COM呼出しを伴う制御用破棄終了を必須とせず、手動復旧を使います。
未対応の操作・引数を黙って無視しません。初期ライフサイクルとPoC-1の提供範囲は以下に示します。残りのPoCの提供範囲は各操作の実装前に確定します。

## 初期ライフサイクルの提供範囲

R1〜R5の採用契約です。実装・実機検証の完了を示しません。後述のP0-core最終契約に対する初期の制限です。
記録形式と遷移は[設計](architecture.md#initial-lifecycle)を正本とします。

- `session_end`の`discard`と、`session_start`・`session_end`の`force`は省略時をfalseとします。明示的なfalseも通常経路です。trueは副作用の前に `INVALID_REQUEST`、`detail: NOT_IMPLEMENTED` で拒否します。
- 通常実行中の競合は `CONFLICT`、`detail: OPERATION_BUSY` です。isolatedでの公開終了は `CONFLICT`、`detail: SESSION_ISOLATED` とします。未対応のtrueは先に拒否し、対応する要求ではbusyの競合判定を先に行います。
- 開始中の状態照会は `closed`、終了中は開始元の状態を返し、いずれも `busy: true` とします。COMを使わず保持情報から応答します。
- 結果不明は `STATE_UNKNOWN` とisolatedを返します。期限後の成功で応答を成功へ変更しません。後続工程のジョブも同じ規則に従います。
- EOFはAPI応答先がなければ診断に結果を残します。公開の破棄終了を内部的に呼べると仮定しません。

| 条件 | 共通コード / detail | 状態・診断 |
| --- | --- | --- |
| 所有記録の形式不正 | CONFLICT / SESSION_RECORD_INVALID | 生成前はclosed。開始拒否 |
| 記録の読取り・更新・確認のI/O失敗 | INTERNAL_ERROR / SESSION_RECORD_IO_ERROR | 生成前はclosed、生成着手後はisolated。失敗段階とExcel終了確認の成否を分ける |
| 終了安全確認による拒否 | CONFLICT / EXCEL_NEEDS_REVIEW | 元がreadyの場合だけ確定した管理外ブック等の拒否でreadyを維持。元がisolatedなら隔離維持。同一性不一致・確認不能はisolated |
| 終了安全確認や操作の結果不明 | STATE_UNKNOWN / EXCEL_NEEDS_REVIEW（安全確認時） | isolated。期限切れ・残存可能性を示す |

共有違反による排他取得失敗はCONFLICTです。権限・読取り等のI/O失敗を競合にまとめません。
診断には対象ブック、取得できたプロセス同一性、失敗段階、結果不明の操作、残存可能性、ログ参照先を含めます。
6操作の公開応答のJSON Schemaとenvelopeは、[PoC-1の提供範囲](#poc-1の提供範囲)で定めます。共通コード6種類は増やしません。

## PoC-1の提供範囲

T16で2026-11-01にユーザーが承認した契約です。製品ツールの実装・検証の完了を示しません。
対象は`session_start`・`session_status`・`range_get`・`range_set`・`session_save`・`session_end`です。
本節にない事項は、本書の他の節と[設計](architecture.md)に従います。

### 提供範囲

- セル操作は単一セルだけです。矩形はP0-coreで提供します。
- 単一セルでも、値と数式は`[[x]]`の2次元配列で受け渡します。矩形へ広げるときに形を変えないためです。
- `range_get`の`fields`は`values`と`formulas`を提供します。`merged_ranges`は`INVALID_REQUEST`、`detail: NOT_IMPLEMENTED`とします。
- 結合セルへの書込み制限は、PoC-1でも契約どおりに適用します。
- マクロ・テストのジョブはありません。`session_status`はジョブの欄を持ちません。

### 応答の形

成功は`isError: false`とし、`structuredContent`に次の形のオブジェクトを返します。同じJSONをテキストでも返します。

```json
{ "session": { "state": "ready", "busy": false }, "workbook": "<主ブックの絶対パス>" }
```

`session`は応答を確定した時点の状態です。`session`以外のフィールドは操作ごとに決めます。
`outputSchema`は成功の形だけを記述します。ルートには`additionalProperties`を指定せず、後続の項目追加を互換に保ちます。

失敗は`isError: true`とし、`structuredContent`を付けません。
テキストは次の形の1行のJSONで、`McpException`の文言として返します。
選んだ理由は[ADR 0005](adr/0005-error-as-json-text.md)にあります。

```json
{
  "error": {
    "code": "CONFLICT",
    "detail": "OPERATION_BUSY",
    "message": "Another operation is running. Check session_status and retry after busy is false.",
    "operation": "range_set",
    "stage": null,
    "excel_may_remain": null,
    "process": null,
    "log": null
  },
  "session": { "state": "ready", "busy": true }
}
```

| フィールド | 内容 |
| --- | --- |
| `code` | 共通コード6種類のいずれか |
| `detail` | 下記の一覧にあるUPPER_SNAKEの文字列1つ。追加の情報は他のフィールドに入れる |
| `message` | 英語の説明と、次に取る行動。ブックに由来する文字列を含めない |
| `operation` | 失敗した操作の名前 |
| `stage` | 失敗した段階。副作用の前の拒否ではnull |
| `excel_may_remain` | Excel生成に着手した後の失敗では、残存の可能性を真偽値で示す。それ以外はnull |
| `process` | 取得できた所有Excelの`{pid, started_at_utc, name}`。なければnull |
| `log` | `paths.workspace/logs/`の診断ファイルの絶対パス。書いていなければnull |

`stage`の値は`config`・`lock`・`record`・`process_query`・`excel_start`・`open`・`calculate`・`read`・`write`・`verify`・`save`・`safety_check`・`close`・`process_exit`です。
COMの詳細とスタックトレースはログだけに出します。

`message`に含める再試行の指針は次のとおりです。

- `OPERATION_BUSY`は、`session_status`で`busy: false`を確認してから再試行できます。
- `INVALID_REQUEST`は、引数を直して再試行できます。
- `STATE_UNKNOWN`と`SESSION_ISOLATED`は、再送せず[手動復旧](architecture.md#初期の手動復旧)へ進みます。

### 受付の判定順

1. 引数を検証します。型・必須・未知の引数・組合せ・アドレスの文法・配列の形状の不正は`INVALID_REQUEST`です。未対応の`true`・矩形・`merged_ranges`もこの段階で拒否します。
2. busyなら`CONFLICT`、`detail: OPERATION_BUSY`です。
3. 状態を下表で判定します。
4. 操作を実行します。シートの不在など、COMが必要な判定はこの段階で行います。

引数の不正も上記の形で返します。検査の置き場は[実装基準](development.md#mcpツールの実装基準)に従います。
inputSchemaは`additionalProperties: false`とします。

| 操作 | closed | ready | isolated |
| --- | --- | --- | --- |
| `session_status` | 応答 | 応答 | 応答 |
| `session_start` | 実行 | `CONFLICT` / `SESSION_ACTIVE` | `CONFLICT` / `SESSION_ACTIVE` |
| `range_get`・`range_set`・`session_save`・`session_end` | `CONFLICT` / `SESSION_NOT_STARTED` | 実行 | `CONFLICT` / `SESSION_ISOLATED` |

表はbusyでない場合です。`session_status`はbusyでも応答します。

### 各操作

`fields`の既定値と、`force`・`discard`の省略をfalseとする規則は、前節までの規定どおりです。
アドレスは`$`の有無と大文字小文字を問わず受け付け、`$`なしの大文字へ正規化して返します。
列はXFDまで、行は1048576までです。
`sheet`は大文字小文字を区別して照合します。

| 操作 | 成功の出力（`session`以外） | 注釈 RO / D / I / OW |
| --- | --- | --- |
| `session_start` | `workbook` | false / false / false / false |
| `session_status` | `workbook`（未確定ならnull）、`last_error`（なければnull） | true / false / true / false |
| `range_get` | `sheet`、`address`、要求した`values`・`formulas` | false / false / true / false |
| `range_set` | `sheet`、`address` | false / true / true / false |
| `session_save` | `workbook` | false / true / true / false |
| `session_end` | `workbook`（`state`はclosed） | false / true / false / false |

注釈はreadOnly・destructive・idempotent・openWorldの順です。
`range_get`は`Application.Calculate()`でブックの計算結果を変え、それが保存で残り得るため、readOnlyをfalseとします。
`session_save`はディスク上のファイルを上書きするため、destructiveをtrueとします。

`session_status`の`workbook`は、開始処理で確定した主ブックの絶対パスです。
`last_error`は、受付後に実行した操作の最新の失敗で、`error`と同じ形です。
副作用の前の拒否（引数・未対応・busy・状態）は記録しません。
後続の操作の成功では消さず、次の`session_start`の成功で消します。

`session_start`の`project_path`は、プロジェクトのルートの絶対パスです。
`range_set`は`values`と`formulas`のどちらか1つを要求します。

### 失敗の対応

各行の状態は、[初期ライフサイクル契約](architecture.md#initial-lifecycle)の遷移に従います。

| 操作 | 条件 | code / detail | 状態 |
| --- | --- | --- | --- |
| `session_start` | 相対パス | `INVALID_REQUEST` / `INVALID_ARGUMENT` | 変更しない |
| `session_start` | 設定ファイルがない | `NOT_FOUND` / `CONFIG_NOT_FOUND` | closed |
| `session_start` | 設定の構文・キー・値・拡張子の不正、`workspace`・`dist`の重なり | `INVALID_REQUEST` / `INVALID_CONFIG` | closed |
| `session_start` | 主ブック・VBAソース・テストのディレクトリがない | `NOT_FOUND` / `PATH_NOT_FOUND` | closed |
| `session_start` | 主ブックが読取専用 | `CONFLICT` / `WORKBOOK_READ_ONLY` | closed |
| `session_start` | 主ブックを他のプロセスが使用中 | `CONFLICT` / `WORKBOOK_LOCKED` | closed |
| `session_start` | `.xlwright.lock`の排他を共有違反で取得できない | `CONFLICT` / `PROJECT_LOCKED` | closed |
| `session_start` | 所有記録がstarting、残存、同一性不一致 | `CONFLICT` / `EXCEL_NEEDS_REVIEW` | closed |
| `session_start` | プロセスの照会に失敗 | `INTERNAL_ERROR` / `PROCESS_QUERY_FAILED` | closed |
| `session_start` | Excelの生成・同一性の取得に失敗 | `EXCEL_ERROR` / `EXCEL_START_FAILED` | isolated |
| `session_start` | 主ブックのOpenが明確に失敗 | `EXCEL_ERROR` / `WORKBOOK_OPEN_FAILED` | isolated |
| `range_get`・`range_set` | シートがない | `NOT_FOUND` / `SHEET_NOT_FOUND` | ready |
| `range_set` | `values`と`formulas`の両方・どちらもない、型の不正、エラー値、`=`で始まらない数式 | `INVALID_REQUEST` / `INVALID_ARGUMENT` | 変更しない |
| `range_set` | 配列の形状がアドレスと一致しない | `INVALID_REQUEST` / `SHAPE_MISMATCH` | 変更しない |
| `range_set` | 結合範囲の左上以外 | `EXCEL_ERROR` / `MERGED_CELL` | ready |
| `range_set` | 保護されたセル | `EXCEL_ERROR` / `CELL_PROTECTED` | ready |
| `range_set` | 読戻しで型と値を確認できない。数式は`HasFormula`で確認する | `EXCEL_ERROR` / `WRITE_NOT_VERIFIED` | ready |
| `range_get`・`range_set` | COM呼出しの明確な失敗 | `EXCEL_ERROR` / `EXCEL_REJECTED` | ready |
| `session_save`・`session_end` | 保存の明確な失敗 | `EXCEL_ERROR` / `SAVE_FAILED` | ready |
| `session_end` | 終了安全確認による拒否。`stage: safety_check`は保存済みを示す | `CONFLICT` / `EXCEL_NEEDS_REVIEW` | 上記の契約による |
| `session_end` | 終了安全確認の結果不明 | `STATE_UNKNOWN` / `EXCEL_NEEDS_REVIEW` | isolated |
| 全操作（`session_status`を除く） | 期限切れ・結果不明 | `STATE_UNKNOWN` / `OPERATION_TIMEOUT` | isolated |
| 全操作（`session_status`を除く） | Excelとの接続が切れた | `STATE_UNKNOWN` / `EXCEL_DISCONNECTED` | isolated |

所有記録の形式不正とI/O失敗は、[初期ライフサイクルの提供範囲](#初期ライフサイクルの提供範囲)の表に従います。
想定外の内部の失敗は`INTERNAL_ERROR`、`detail: INTERNAL_FAILURE`とします。Excel生成に着手した後なら、結果不明と同じく隔離します。

PoC-1で使う`detail`は次のとおりです。

| code | detail |
| --- | --- |
| `INVALID_REQUEST` | `INVALID_ARGUMENT`、`INVALID_ADDRESS`、`SHAPE_MISMATCH`、`INVALID_CONFIG`、`NOT_IMPLEMENTED`、`RESPONSE_TOO_LARGE` |
| `NOT_FOUND` | `CONFIG_NOT_FOUND`、`PATH_NOT_FOUND`、`SHEET_NOT_FOUND` |
| `CONFLICT` | `OPERATION_BUSY`、`SESSION_ACTIVE`、`SESSION_NOT_STARTED`、`SESSION_ISOLATED`、`PROJECT_LOCKED`、`WORKBOOK_LOCKED`、`WORKBOOK_READ_ONLY`、`SESSION_RECORD_INVALID`、`EXCEL_NEEDS_REVIEW` |
| `EXCEL_ERROR` | `EXCEL_START_FAILED`、`WORKBOOK_OPEN_FAILED`、`SAVE_FAILED`、`MERGED_CELL`、`CELL_PROTECTED`、`EXCEL_REJECTED`、`WRITE_NOT_VERIFIED` |
| `STATE_UNKNOWN` | `OPERATION_TIMEOUT`、`EXCEL_DISCONNECTED`、`EXCEL_NEEDS_REVIEW` |
| `INTERNAL_ERROR` | `SESSION_RECORD_IO_ERROR`、`PROCESS_QUERY_FAILED`、`INTERNAL_FAILURE` |

アドレスが矩形なら`NOT_IMPLEMENTED`、不連続・行全体・列全体・文法の不正なら`INVALID_ADDRESS`です。

### JSON Schema

記法の都合で共通の定義を`$defs`にまとめます。各ツールのSchemaでは展開してもかまいません。
方言はMCPの既定の2020-12です。

```json
{
  "$defs": {
    "SessionSummary": {
      "type": "object",
      "properties": {
        "state": { "enum": ["closed", "ready", "isolated"] },
        "busy": { "type": "boolean" }
      },
      "required": ["state", "busy"]
    },
    "CellError": {
      "type": "object",
      "properties": {
        "type": { "const": "error" },
        "code": {
          "enum": ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!", null]
        },
        "raw": { "type": "integer" }
      },
      "required": ["type", "code"]
    },
    "CellValue": {
      "anyOf": [{ "type": ["number", "string", "boolean", "null"] }, { "$ref": "#/$defs/CellError" }]
    },
    "WritableValue": { "type": ["number", "string", "boolean", "null"] },
    "Formula": { "type": ["string", "null"] },
    "Error": {
      "type": "object",
      "properties": {
        "code": { "enum": ["INVALID_REQUEST", "NOT_FOUND", "CONFLICT", "EXCEL_ERROR", "STATE_UNKNOWN", "INTERNAL_ERROR"] },
        "detail": { "type": "string" },
        "message": { "type": "string" },
        "operation": { "type": "string" },
        "stage": { "type": ["string", "null"] },
        "excel_may_remain": { "type": ["boolean", "null"] },
        "process": {
          "type": ["object", "null"],
          "properties": {
            "pid": { "type": "integer" },
            "started_at_utc": { "type": "string" },
            "name": { "type": "string" }
          },
          "required": ["pid", "started_at_utc", "name"]
        },
        "log": { "type": ["string", "null"] }
      },
      "required": ["code", "detail", "message", "operation", "stage", "excel_may_remain", "process", "log"]
    }
  }
}
```

`code`が`null`の`CellError`には`raw`を付けます。

| 操作 | inputSchema |
| --- | --- |
| `session_start` | `{"type":"object","properties":{"project_path":{"type":"string"},"force":{"type":"boolean"}},"required":["project_path"],"additionalProperties":false}` |
| `session_status`・`session_save` | `{"type":"object","additionalProperties":false}` |
| `range_get` | `{"type":"object","properties":{"sheet":{"type":"string"},"address":{"type":"string"},"fields":{"type":"array","items":{"enum":["values","formulas","merged_ranges"]},"minItems":1,"uniqueItems":true}},"required":["sheet","address"],"additionalProperties":false}` |
| `range_set` | `{"type":"object","properties":{"sheet":{"type":"string"},"address":{"type":"string"},"values":{"type":"array","items":{"type":"array","items":{"$ref":"#/$defs/WritableValue"}}},"formulas":{"type":"array","items":{"type":"array","items":{"$ref":"#/$defs/Formula"}}}},"required":["sheet","address"],"additionalProperties":false}` |
| `session_end` | `{"type":"object","properties":{"discard":{"type":"boolean"},"force":{"type":"boolean"}},"additionalProperties":false}` |

| 操作 | outputSchema |
| --- | --- |
| `session_start`・`session_save`・`session_end` | `{"type":"object","properties":{"session":{"$ref":"#/$defs/SessionSummary"},"workbook":{"type":"string"}},"required":["session","workbook"]}` |
| `session_status` | `{"type":"object","properties":{"session":{"$ref":"#/$defs/SessionSummary"},"workbook":{"type":["string","null"]},"last_error":{"anyOf":[{"$ref":"#/$defs/Error"},{"type":"null"}]}},"required":["session","workbook","last_error"]}` |
| `range_get` | `{"type":"object","properties":{"session":{"$ref":"#/$defs/SessionSummary"},"sheet":{"type":"string"},"address":{"type":"string"},"values":{"type":"array","items":{"type":"array","items":{"$ref":"#/$defs/CellValue"}}},"formulas":{"type":"array","items":{"type":"array","items":{"$ref":"#/$defs/Formula"}}}},"required":["session","sheet","address"]}` |
| `range_set` | `{"type":"object","properties":{"session":{"$ref":"#/$defs/SessionSummary"},"sheet":{"type":"string"},"address":{"type":"string"}},"required":["session","sheet","address"]}` |

アドレスの文法、配列の形状、`values`・`formulas`の排他、数式の先頭の`=`は、Schemaではなくツール側で検査します。
ツールの説明は英語で、用途・副作用・失敗後に再試行できるかを書きます。文面は実装時にスナップショットの差分で確認します。

### 未確認

- `McpException`の文言がそのままテキストになることは、1箇所でだけ確認しました。[調査記録](notes/mcp-best-practices.md#対象の版と実測)を参照してください。
- 引数をツール側で検査する形を、SDK 2.3.0で組めるかは未確認です。
- 単一セルでの計算鮮度の手順と読戻しの照合は、製品経路で未検証です。

## 共通規約

- セッションはサーバープロセスごとに最大1つとし、MCPのプロトコルのセッションに依存しません。`session_id` は引数に取りません。理由は[ADR 0004](adr/0004-process-scoped-session.md)にあります。
- `project_path` は正規化可能な絶対パスです。操作対象は開始時に確定した主ブックです。
- シート名と範囲は明示し、ActiveWorkbook・ActiveSheetに依存しません。
- 通常の操作は結果を直接返し、マクロとテストは `execution_id` を返します。
- 実行中の競合操作は `CONFLICT`、`detail: OPERATION_BUSY` とします。状態照会はCOMを使わず受け付けます。
- 応答上限は1MiBです。超える場合は切り詰めず `INVALID_REQUEST`、`detail: RESPONSE_TOO_LARGE` とします。
- エラーには説明、セッション状態、`detail`、必要に応じてログ参照先を含めます。

共通エラーコードは `INVALID_REQUEST`・`NOT_FOUND`・`CONFLICT`・`EXCEL_ERROR`・`STATE_UNKNOWN`・`INTERNAL_ERROR` の6種類です。
入力不正や明確な不在はセッションを隔離しません。
結果不明の場合は `STATE_UNKNOWN` とし、セッションを `isolated` にします。
遅延した成功で応答やジョブ結果を成功へ書き換えず、通常の操作列と保存を再開しません。
未実装のP0-ext操作・フィールドは `INVALID_REQUEST`、`detail: NOT_IMPLEMENTED` とします。

## 操作一覧

`?` は省略可能を表します。省略時の挙動が本書で決まっていないものは実装前に確定します。

| 操作 | 入力 | 結果・用途 |
| --- | --- | --- |
| `workbook_create` | `project_path` | テンプレートからブック作成。Excelを起動せず、既存ファイルは上書きしない |
| `project_init` | `project_path`, `existing_workbook_path` | 既存ブックのコピーと初回VBA抽出 |
| `session_start` | `project_path`, `force?` | 設定検証、排他、残存Excelの後始末、主ブックを開く |
| `session_status` | なし | `state`, `busy`, ブック、ジョブ、直近のエラーとログ参照先 |
| `session_save` | なし | 主ブックを保存し、セッションを維持 |
| `session_end` | `discard?`, `force?` | 通常は保存して終了。`discard` で保存せず終了 |
| `range_get` | `sheet`, `address`, `fields?` | 値・数式・交差する結合範囲を取得 |
| `range_set` | `sheet`, `address`, `values` または `formulas` | 矩形へ値または数式を書き込む |
| `worksheet_list` | なし | 名前、位置、表示状態 |
| `worksheet_create` | `name` | シート作成 |
| `worksheet_delete` | `name` | シート削除 |
| `worksheet_rename` | `old_name`, `new_name` | シート改名 |
| `vba_module_list` | なし | 名前、種類、対応ソース、管理対象かどうか |
| `vba_sync_sources` | `source_paths?` | ソースをブックへ反映 |
| `vba_module_delete` | `name` | 標準・クラスモジュールの明示削除 |
| `macro_run` | `name`, `args?`, `dialog_responses?`, `timeout_seconds?` | 非同期マクロ実行 |
| `test_list` | なし | 注釈で登録されたVBA内部テストの一覧 |
| `test_run` | `selection`, `dialog_responses?`, `startup_events?`, `timeout_seconds?` | 非同期テスト実行 |
| `execution_status` | `execution_id` | 実行状態と、完了時の結果 |

`session_start` の `force` は、正しい所有記録で識別できた残存Excelに限ります。
壊れた所有記録や他プロジェクトのロックを突破できません。
`session_end` の `force` は `discard: true` の場合だけ受け付けます。
人の承認等の条件は[終了と後始末](architecture.md#終了と残存excelの後始末)を参照してください。

`project_init` は取り込み元を変更しません。
XLMマクロシートがあれば `INVALID_REQUEST`、`detail: XLM_MACRO_PRESENT` とします。
警告やVBIDEのアクセス拒否を自動承認・迂回しません。
初期化に失敗した途中生成物は削除せず、成功とも報告しません。

## セルと範囲

`sheet` は正確なシート名です。存在しなければ `NOT_FOUND` とします。
`address` はシート名を含まないA1形式の単一セルまたは矩形です。`$` 付きも許容します。
不連続範囲、行全体、列全体は拒否します。

### 読み取り

`fields` の既定は `["values"]` です。
`values`・`formulas`・`merged_ranges` を指定できます。書式はP0-extです。

値と数式は、単一セルでも2次元配列で返します。
値は数値・文字列・真偽値・空白の `null` です。日付はシリアル値で返します。
Excelエラーは `{"type":"error","code":"#DIV/0!"}` 等のオブジェクトにします。
変換できないエラーは `{"type":"error","code":null,"raw":<整数>}` とします。
表示文字列は返しません。

値は `Value2`、数式は `HasFormula` が真のセルの `Formula2` から取得する設計です。
数式がないセルの `formulas` は `null` とします。
S3の固定ケースではスピル子のHasFormula=false、S4では従来7種とSPILL・CALCのValue2がInt32で戻ることを確認しました。
検証範囲と保存検査の補正は[T8の結果](notes/spike-0-results.md#t8の非com確認2026-10-30)を参照してください。製品APIの実装・全エラー種の保証は含みません。
結合範囲の値は左上だけに入り、他の位置は `null` とします。
`merged_ranges` は取得範囲と交差する結合領域をA1文字列の配列で返します。

#### 計算結果の鮮度

S14の結果と2026-10-31のユーザー承認に基づく設計契約です。
製品APIの実装・実機検証の完了を示しません。

`fields`に`values`を含む`range_get`は、`Application.Calculate()`、`CalculationState=xlDone`の確認、`Value2`取得の順に実行します。
自動・手動で取得手順を分岐せず、計算モードを一律に変更しません。
マクロが計算モードを変えた場合も、この取得境界で再計算・完了確認を行います。
再計算・完了確認・値取得は、同じ同期操作の期限に含めます。
期限超過や結果不明は共通規約に従い、古い値を成功として返しません。

対象は単一の管理ブック内の同期計算です。
外部データ更新・非同期UDF・揮発性や反復計算・任意VBA・複数ブックの鮮度は保証しません。
S14の実測は2つの算術式と既知3マクロに限ります。
Calculate復帰後に追加待機が必要な遷移、照会と値取得の原子性は未実証です。
詳細は[S14の結果と採用判断](notes/spike-0-results.md#t9の完了判定と採用判断)を参照してください。

### 書き込み

`values` と `formulas` の同時指定は拒否します。
単一セルでも2次元配列を要求し、範囲と形状が一致しなければ変更前に `INVALID_REQUEST` とします。

| 入力 | 契約 |
| --- | --- |
| `null` | 値と数式を消去し、書式と結合を維持 |
| `""` | 長さ0の文字列として格納 |
| `values` の文字列 | 数値・日付・数式へ解釈せず、文字列として格納 |
| `formulas` の非null要素 | `=` で始まる英語表記・A1形式の数式だけを許可 |
| 数値・論理値 | その型の値として格納 |
| エラー値オブジェクト | 書き込み不可 |

文字列は先頭にアポストロフィを付けて書き込む候補方式を、S1の固定ケースで評価しました。空文字を含む文字列保持の仮説を支持しました。
この方式では `PrefixCharacter` がアポストロフィになります。
`null` の一括書き込みはS2の固定ケースで、値・数式の消去と採取した書式属性・結合の維持を確認しました。個別ClearContentsも対照として確認しました。
任意の書式・結合配置の網羅ではありません。必要なら対象セルへの個別ClearContentsを検討します。
書き込み後は読み戻して型と値を照合します。

結合範囲の左上1セルへの書き込みだけを許可します。
左上以外、または結合セルを含む複数セルへの書き込みは拒否します。
結合解除、書き込み先の自動補正、保護解除は行いません。
保護・結合の制約違反は `EXCEL_ERROR` とし、理由を `detail` に含めます。
途中まで変更された失敗を成功と報告せず、分かる範囲の処理済み部分を返します。

## VBAの反映

- `source_paths` は `paths.vba_source` 配下の既存ファイルに限ります。開始時に対象内容を確定します。
- ソースはUTF-8とし、BOM付きも許容します。Import前にCRLFとシステムANSIコードページへ変換する案をS5で検証します。
- 変換不能なら変更前に `INVALID_REQUEST`、`detail: VBA_ENCODING` とし、ファイル名と行番号を含めます。
- 標準・クラスは、既存の一時改名 → Import → 旧モジュールのRemoveをS6で検証します。失敗時は元の名前への復帰を試みますが、自動ロールバックは保証しません。
- `.cls` のヘッダー欠落はエラーです。ドキュメントモジュールは全行置換し、先頭のモジュール属性行を除去します。
- ドキュメントモジュールのプロシージャ属性は `INVALID_REQUEST`、`detail: UNSUPPORTED_ATTRIBUTE` とします。
- ドキュメントモジュールは削除・再作成しません。

`macro_run` と `test_run` は実行前にソースを反映します。
セッションの最初の実行は全管理対象モジュールを反映し、その後はソースのハッシュが変わったモジュールだけを反映します。
反映状態は永続化しません。VBEの読み戻し本文による同期判定も行いません。
反映に失敗したら実行を開始しません。

## マクロ・テスト・ダイアログ

以下は既存のラッパー方式を基準とする設計です。
R4のS7・S15では、`Application.Run`と最小限のエラー捕捉、.NET標準機能によるJSON変換で満たせる範囲を比較します。
比較だけで `xw_Runtime` や公開結果形式を変更・削除せず、採用する契約を確定してから実装します。
R5ではマクロ実行に必要な部分を扱い、ダイアログラッパーとテストハーネスはR6で追加します。

`xw_Runtime` は標準モジュール1つで、実行ラッパー、ダイアログラッパー、テストハーネスを提供します。
初期化時にソースへ配置し、その後は管理対象ソースとして扱います。更新は手作業のコピーです。
存在しなければ `CONFLICT`、`detail: RUNTIME_MISSING` とします。

`macro_run` の引数は数値・文字列・論理値・`null` に限定します。
`Application.Run` で `xw_Runtime.Xw_Run` を呼び、成功・失敗、戻り値、VBAエラー、ダイアログ履歴を固定形式のJSON文字列で受け取る設計です。
捕捉した実行時エラーはマクロの失敗とし、セッションは `ready` を維持します。
コンパイルエラーや捕捉できない停止の挙動はS7で検証します。

テストは標準モジュールの宣言部の `'@TestModule` と、引数なし `Public Sub` の直前行の `'@TestMethod` で登録します。
IDは `<モジュール名>.<プロシージャ名>` です。名前から推測しません。
`selection` はIDの配列または `all` です。初期化・後始末用の注釈はP0では設けません。
`Xw_AssertEqual`・`Xw_AssertTrue`・`Xw_Fail` の専用エラー番号により `passed`・`failed`・`error` を区別します。

ダイアログラッパーは `Xw_MsgBox`・`Xw_InputBox`・`Xw_GetOpenFilename`・`Xw_SelectFolder`・`Xw_GetSaveAsFilename` です。
`dialog_responses` は呼出し順の期待する種類と応答（値、パス、キャンセル）の配列です。
種類は `msgbox`・`inputbox`・`open_file`・`select_folder`・`save_file` とします。
種類や順序の不一致、応答の不足は既定値を返さず失敗とします。
キュー未設定の通常実行では標準ダイアログを表示します。実行終了時は必ずキューを消去します。
ラッパー外のダイアログは自動操作しません。

`execution_status` は `running`・`completed`・`failed`・`timed_out`・`state_unknown` を返します。
完了時はマクロ結果またはテストケースごとの結果、エラー、ダイアログ履歴、ログ参照先も含めます。
タイムアウトを処理停止の証明として返しません。

## 実装前に確定する詳細

- PoC-1の6操作以外のJSON Schemaと`detail`。envelopeはPoC-1の形を使います。
- 矩形と`merged_ranges`を提供するときの、処理済み部分の返し方。
- PoC-2以降で提供する操作・引数の範囲と、未対応要求の拒否応答。PoC-1の分は確定済みです。
- 省略可能引数の既定値と型、`startup_events` の指定方法。
- `isolated` と保留中呼出しが同時に存在する場合の `session_end` の受付条件。初期は上記のとおり拒否し、後続の受付条件は[安全条件の承認](architecture.md#承認待ちの安全条件)後に反映します。
- ラッパー経由のマクロ引数の有効上限と、戻り値の対応範囲。
- ダイアログ応答、テスト結果、実行結果JSONの具体的なスキーマ。
- `xw_Runtime` の版表現と対応確認。
- セル・数式・VBAの仮説が外れた場合の契約変更。
- 計算鮮度の上記設計を製品経路で検証すること。S14の実測範囲外を拡張する場合の保証と対応範囲。
- ラッパーと対象マクロの実行先を、主ブックまたは現在のテストコピーに限定する名前解決（S15で確認）。
- ソースファイルとVBAコンポーネントの対応。特にドキュメントモジュールとシート名・CodeNameの区別。
- 起動イベント実行前にダイアログ応答を準備する方法と、未準備時の扱い。`startup_events` を実装する前に確定します。
