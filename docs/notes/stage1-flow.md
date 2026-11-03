# 段階1の流れを会話から通す

[ROADMAP](../../ROADMAP.md)の段階1の完了の根拠を確かめた記録です。Claude Codeの会話から、開発版のプラグインのツールだけで流れを通すことと、終わらないマクロで実行時間の上限の終了を確かめることの2つです。
確かめた日は2026-11-03です。

## 版

- Windows 11 Pro 10.0.26300、Claude Codeのデスクトップアプリ（Codeタブ）、auto mode
- ビルド: `6ba3037-dirty-20261103T120611Z`（`server_info`の`build_id`と`.xlwright/builds/latest.json`が一致）
- サーバーのコマンドラインの`--project`は、リポジトリの完全なパス（`Win32_Process`で照会）
- Microsoft 365 Excel 64-bit 16.0.20527.20112。`AccessVBOM`は1、`VBAWarnings`は2
- 事前の許可: `.claude/settings.local.json`（Git管理外）の`permissions.allow`に、`mcp__plugin_xlwright-dev_xlwright__`の8つのツールを書いた

## 再現の手順

1. 捨てる前提のアプリのフォルダ`.xlwright/scratch/stage1/`（Git管理外）を作ります。主ブック`Stage1.xlsm`は、テストの`MinimalWorkbook`と同じ最小の形で、シートは`Sheet1`だけです。`src/modules/Calc.bas`に、次の2つを置きます。
   - `Add`: `Sheet1`の`A1`と`A2`を足して`B1`に書き、`"合計は " & B1`を返す`Function`。
   - `Forever`: 中身の無い`Do`〜`Loop`だけの`Sub`。
2. 実行の前に、`EXCEL.EXE`のプロセスの数と主ブックのハッシュを照会します。
3. 会話から、次の順にツールを呼びます。`session_start`、`range_write`（`A1:A2`に2と3）、`vba_sync`、`macro_run`（`Calc.Add`）、`range_read`（`A1:B2`）、`workbook_save`、`session_end`（`save`が`true`）。
4. 終了の後に、専用のExcelのPIDと`EXCEL.EXE`の数と主ブックのハッシュを照会します。
5. `session_start`で開き直して`range_read`で保存の中身を読み、`macro_run`で`Calc.Forever`を`timeout_seconds`が5で実行します。終了の後に、手順4と同じ照会をします。

## 判定

### 流れ

支持。7つの呼出しは、どれも人の操作なしで成功した。

- `macro_run`は`"合計は 5"`を返した（`elapsed_ms`は102）。`sync`には`Calc`と`XlWright`が出た。
- `range_read`は`[[2,5],[3,null]]`を返した。
- `session_end`の後、専用のExcel（PID 13352）は残らなかった。`EXCEL.EXE`の数は、実行の前と同じ0だった。
- 主ブックのハッシュと更新の日時は変わった。開き直した`range_read`も`[[2,5],[3,null]]`を返した。

### 実行時間の上限

支持。`Calc.Forever`は、`code`が`STATE_UNKNOWN`、`detail`が`TIMEOUT`、`data.excel_ended`が`true`で返った。`session`は`closed`になった。
専用のExcel（PID 8452）は残らず、`EXCEL.EXE`の数は0だった。主ブックのハッシュは、保存の後のままだった。

### 実行の間の`session_status`

不一致（1回）。`macro_run`（`Calc.Forever`、`timeout_seconds`が8）と`session_status`を、会話の1つのメッセージで並べて呼んだ。`session_status`は`closed`を返した。
`macro_run`が終わった後に`session_status`が実行されたと見られる。この形では、実行の間の状態を会話から照会できなかった。専用のExcel（PID 19124）は残らなかった。

### 事前の許可

結果不明。会話がauto modeで動いていたため、許可の行が無い時に確認の画面が出るかを見分けられなかった。

## 後始末とGit差分

- `EXCEL.EXE`は、すべての照会で0だった。
- `.xlwright/scratch/stage1/.xlwright/sync/`が、空のフォルダとして残った。
- 実験の後の`git status`に、差分は無かった。

## 未確認

- auto modeでない会話で、`.claude/settings.local.json`の許可だけで確認の画面が出ないこと。
- 会話から、実行の間に`session_status`を照会する方法。
