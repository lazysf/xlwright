# 段階4の操作を会話から通す

[ROADMAP](../../ROADMAP.md)の段階4の完了の根拠のうち、「この段階で加える操作を、実Excelのシナリオで確かめる」を確かめた記録です。新しいビルドを読み込んだ別のセッションの会話から、開発版のツールとスキルだけで、6つのシナリオを人の操作なしで通しました。確かめた日は2026-11-05です。
「自分のアプリ1つで、依頼から報告まで回す」は、利用する版で行うため、配り方（[記録](distribution.md)）と一緒に保留しています。

## 版

- Windows 11 Pro 10.0.26300、Claude Codeのデスクトップアプリ（Codeタブ）
- ビルド: `fba2364-dirty-20261105T095617Z`（会話の`server_info`。`latest.json`と一致）
- Microsoft 365 Excel 64-bit 16.0.20527.20112。`AccessVBOM`は1、`VBAWarnings`は2
- 開発版のプラグインのスキルは、段階4のツールの扱いを加えた版（人が承認、2026-11-05）

## 再現の手順

- 捨てる前提のアプリのフォルダを、`.xlwright/scratch/stage4/s1`〜`s6`（Git管理外）にシナリオごとに作ります。材料も、会話がツールで作ります。
- 1の既存のブックは、`workbook_create`で作った`Legacy.xlsm`に、`Workbook_Open`がアプリのフォルダの`opened.txt`へ追記するコードと標準モジュールを入れて保存し、`session_end`の後に`src/`を消して用意します。
- 2の`.xlsx`も`workbook_create`で作ります。
- 前後で`EXCEL.EXE`の数を照会します。

## 判定

### 1. 既存の`.xlsm`から`vba_extract`で取り出す

- 支持。`src/workbook/ThisWorkbook.cls`・`src/workbook/Sheet1.cls`・`src/modules/Legacy.bas`が書き出され、`skipped`は空、`XlWright`のモジュールは書き出されなかった。取り出したソースを反映した`macro_run`の`Legacy.Twice(21)`は42を返した。
- `opened.txt`は、取り出しの間も後も無かった。ただし`session_start`で主ブックを開いた時にも`Workbook_Open`は動かない（開く時にイベントを止めるため）。取り出しだけで動かないことは、この会話でも実Excelの結合テストでも切り分けていない。イベントを止めずに`ForceDisable`だけで開いた写しで動かないことは、周期38の実験が確かめている（[記録](vba-extract.md)）。印を書くコード自体が動くことは、この会話では確かめていない。

### 2. `.xlsx`を`.xlsm`へ変える

- 支持。変える前の`workbook_get`は`format`が`xlsx`、`saved`が真。`convert_to_xlsm`は`workbook`が`Book.xlsm`、`previous`が`Book.xlsx`を返し、その後の`workbook_get`は`xlsm`で、A1の値も残っていた。
- 次の`session_start`は`MULTIPLE_WORKBOOKS`（`workbooks`は`Book.xlsm`と`Book.xlsx`）で止まり、`.xlsx`を消すと`Book.xlsm`で開始できた。

### 3. 空のアプリのフォルダで主ブックを作る

- 支持。開始の`workbook`は`null`、作る前の`workbook_get`は`NO_WORKBOOK`で`workbook_create`を案内した。`workbook_create`の後の`workbook_get`は`xlsm`、`saved`が真。

### 4. 参照設定を加える

- 支持。加える前の`vba_check`は`compiled`が偽（ユーザー定義型が定義されていない）。`vba_edit`の`add_reference`（Scripting Runtime 1.0）は`changed`が真で、その後の`vba_check`は通り、`Scripting.Dictionary`を使う`macro_run`は値を返した。保存して開き直した後の同じ追加は`changed`が偽で、参照設定は保存の後も残った。

### 5. ダイアログに事前に答える

- 支持。`macro_run`の`dialogs`に`message`の`yes`と`input`の値を渡すと、マクロは渡した値を使い、`dialogs`は`answered`が2、`unused`が0。答えが余ると`unused`が1、足りないと`MACRO_ERROR`（`source`が`XlWright`、`XlWright.InputText: no answer is left for the dialog.`）で、セッションは続いた。
- `XlWright.Answer`で答えるテスト2件（答える場合と`False`で取り消す場合）は、`test_run`で通った。

### 6. まだ無い関数を呼ぶテストを先に書く

- 支持。最初の`test_run`は`COMPILE_ERROR`（`data.error`はテストのモジュールの8行目、メソッドまたはデータ メンバーが見つからない）で、セッションは`ready`のまま、`EXCEL.EXE`はセッションの1つだけだった。関数を書いた次の`test_run`で2件とも通った。

## 分かったこと

- `workbook_get`と`range_read`を会話から並べて送ると、`range_read`は`BUSY`で拒否された。単独で送り直すと通った。契約どおり（実行中は待たせずに拒否する）。
- `.xlwright/extract/`（s1）と`.xlwright/test/`（s5・s6）が、空のフォルダとして残った。`.xlwright/sync/`の残りは既知（[TODO](../../TODO.md)の書き留め）。

## 後始末とGit差分

- `EXCEL.EXE`は、始める前も終えた後も0。各シナリオの後に`session_end`で終え、所有を確かめられないExcelには触れていない。`STATE_UNKNOWN`は出なかった。
- 会話は、`.xlwright/scratch/stage4/`の外を変えていない（`git status`に変更なし）。材料は`.xlwright/scratch/stage4/`に残してあった（2026-11-10に消した。作り直しは再現の手順のとおり）。
