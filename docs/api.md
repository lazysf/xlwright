# P0-core API契約

本書は実装対象の操作契約です。利用可能な実装を示すものではありません。
MCPのJSON Schemaと応答全体のスキーマは、各操作の実装前に確定します。
未確定のフィールドを、実装上の都合だけで公開仕様にしないでください。
実行制御と安全条件は [設計文書](architecture.md) に従います。

## 共通規約

- セッションはMCP接続ごとに1つとし、`session_id` は引数に取りません。
- `project_path` は正規化可能な絶対パスです。操作対象は開始時に確定した主ブックです。
- シート名と範囲は明示し、ActiveWorkbook・ActiveSheetに依存しません。
- 通常の操作は結果を直接返し、マクロとテストは `execution_id` を返します。
- 実行中の競合操作は `CONFLICT`、`detail: OPERATION_BUSY` とします。状態照会はCOMを使わず受け付けます。
- 応答上限は1MiBです。超える場合は切り詰めず `INVALID_REQUEST`、`detail: RESPONSE_TOO_LARGE` とします。
- エラーには説明、セッション状態、`detail`、必要に応じてログ参照先を含めます。

共通エラーコードは `INVALID_REQUEST`・`NOT_FOUND`・`CONFLICT`・`EXCEL_ERROR`・`STATE_UNKNOWN`・`INTERNAL_ERROR` の6種類です。
入力不正や明確な不在はセッションを隔離しません。
結果不明の場合は `STATE_UNKNOWN` とし、セッションを `isolated` にします。
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
数式がないセルの `formulas` は `null` とします。スピルの挙動とエラー型はS3・S4で検証します。
結合範囲の値は左上だけに入り、他の位置は `null` とします。
`merged_ranges` は取得範囲と交差する結合領域をA1文字列の配列で返します。

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

文字列は先頭にアポストロフィを付けて書き込む案をS1で検証します。
この方式では `PrefixCharacter` がアポストロフィになります。
`null` の一括書き込みはS2で検証し、必要なら対象セルへ個別に `ClearContents` を行います。
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

- 各操作のJSON Schema、出力のenvelope、`detail` の具体的な構造。
- 省略可能引数の既定値と型、`startup_events` の指定方法。
- `isolated` と保留中呼出しが同時に存在する場合の `session_end` の受付条件。[安全条件の承認](architecture.md#承認待ちの安全条件)後に反映します。
- ラッパー経由のマクロ引数の有効上限と、戻り値の対応範囲。
- ダイアログ応答、テスト結果、実行結果JSONの具体的なスキーマ。
- `xw_Runtime` の版表現と対応確認。
- セル・数式・VBAの仮説が外れた場合の契約変更。
- 計算モード、再計算の起点、完了待ち、値取得時に保証する結果の鮮度（S14で確認）。
- ラッパーと対象マクロの実行先を、主ブックまたは現在のテストコピーに限定する名前解決（S15で確認）。
- ソースファイルとVBAコンポーネントの対応。特にドキュメントモジュールとシート名・CodeNameの区別。
- 起動イベント実行前にダイアログ応答を準備する方法と、未準備時の扱い。`startup_events` を実装する前に確定します。
