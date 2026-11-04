# 段階3を会話から通す

[ROADMAP](../../ROADMAP.md)の段階3の完了の根拠を確かめた記録です。小さなブックで、ソースの修正、反映、テストの実行、失敗したテストを直して通す、を人の操作なしで繰り返しました。要件の[起こし方](../requirements.md#確かめ方)のうち、テストの中で起こす2つも通しました。
確かめた日は2026-11-04です。

## 版

- Windows 11 Pro 10.0.26300、Claude Codeのデスクトップアプリ（Codeタブ）、auto mode
- ビルド: `0f7e516-dirty-20261104T114815Z`（会話A・B・Cの`server_info`。`latest.json`と一致）
- Microsoft 365 Excel 64-bit 16.0.20527.20112。`AccessVBOM`は1、`VBAWarnings`は2
- 開発版のプラグインのスキルは、段階3のツールの扱いを加えた版（人が承認、2026-11-04）

## 再現の手順

- 捨てる前提のアプリのフォルダ`.xlwright/scratch/stage3/`（Git管理外）を使います。主ブックは`Stage3.xlsm`（`stage1`の写し）、`data/sample.csv`、`src/modules/Text.bas`（`FieldCount`が1つ少なく数える）、`tests/TextTests.bas`（`FieldCount`の2つのテスト）です。
- アプリのフォルダの外のブックとして、`.xlwright/scratch/outside/Outside.xlsm`（`stage1`の写し）を置きます。
- 元の会話から`start_session`で会話A・B・Cを始め、手順をメッセージで送ります。元の会話のサーバーは古いビルドのため、Excelの操作には使いません。
- 各手順の前後で、`EXCEL.EXE`のプロセスと所有記録を照会します。

## 判定

### 修正・反映・テスト・直して通す（会話A）

支持（往復2回）。人の操作は無かった。

- 往復1: `vba_check`は`compiled`が真。`test_run`は`passed`0・`failed`2で、`message`は「AreEqual: expected 3, actual 2」と「AreEqual: expected 1, actual 0」。会話Aは、テストを変えずに`FieldCount`を`UBound(...) + 1`に直し、`vba_check`と`test_run`で2つとも通した。
- 往復2: 会話Aは、`JoinFields`のテストを先に書いて`test_run`を呼んだ。まだ無い関数を呼ぶテストは、失敗したテストの結果でなく、コンパイルエラーのダイアログ（「メソッドまたはデータ メンバーが見つかりません。」）で`STATE_UNKNOWN`・`DIALOG`になった。`excel_ended`と`restartable`はともに真で、`data.test`は止まったテストだった。会話Aは開始し直し、`JoinFields`を作り、3つとも通した。

### テストの中の実行時エラー・コンパイルエラー・終わらないマクロ（会話A）

支持（各1回）。どれも、`tests`で1つのテストだけを指定した。

- 実行時エラー（0で割る）: 失敗の応答でなく結果で返った。`errors`1、`number`11、`description`は「0 で除算しました。」。セッションは`ready`のままだった。
- コンパイルエラー（`Option Explicit`で宣言の無い変数）: `STATE_UNKNOWN`・`DIALOG`。ダイアログは「変数が定義されていません。」。`excel_ended`と`restartable`はともに真で、開始し直せた。
- 終わらないマクロ（`Do: DoEvents: Loop`、`timeout_seconds`10）: `STATE_UNKNOWN`・`TIMEOUT`。`excel_ended`と`restartable`はともに真で、開始し直せた。
- 各回の後にテストのファイルを消し、`test_run`（全部）は3つとも通った。最後の`session_end`は成功した。

### テストのマクロで外のブックを開いたまま会話を閉じ、次の会話で開始する（会話B・C）

支持（1回）。

- 会話Bは、`Workbooks.Open`で`Outside.xlsm`を開いて閉じないテストを`test_run`で実行した（`passed`1）。`session_end`を呼ばずに終え、元の会話がアーカイブした。サーバーは終わり、専用のExcel（PID 16752）は所有記録の`owned`のまま残った。
- 会話Cの`session_start`は`CONFLICT`・`LEFTOVER_EXCEL`で止まった。`message`は、アプリのフォルダの外のブックが開いているため閉じずに残したことと、人が確かめて閉じることを書いていた。`data`には`record`・`excel`（PID 16752と生成時刻）・`log`が入った。PID 16752は終了されず、所有記録も変わらなかった。
- 人が表示要求のコマンド（`plugins/xlwright-dev/scripts/show-excel.ps1 -App .xlwright\scratch\stage3`）でExcelを見て、2つのブックを保存せずに閉じ、`done`と入力した。コマンドは、ブックのウィンドウが無い形で残ったプロセスを終了させ、終了コード0で終わった。
- その後、会話Cの`session_start`は成功し（`cleanup`は無し）、所有記録は完了になった。`session_end`も成功した。

## 分かったこと

- まだ無い関数を呼ぶテストを先に書くと、`test_run`はコンパイルエラーのダイアログで専用のExcelを終える。`vba_check`を先に呼べば、Excelを終えずにエラーを受け取れる。
- `LEFTOVER_EXCEL`の`message`は、タスク マネージャーで確かめるよう書き、表示要求のコマンドを挙げていない。

## 後始末とGit差分

- 前後の`EXCEL.EXE`は0。所有を確かめられないExcelには触れていない。
- `.xlwright/scratch/stage3/.xlwright/test/`は空。`.xlwright/sync/`は空のフォルダとして残った（[TODO](../../TODO.md)の書き留めにある扱い）。
- Gitの差分は、文書とスキルだけ。管理するブックに触れていない。
