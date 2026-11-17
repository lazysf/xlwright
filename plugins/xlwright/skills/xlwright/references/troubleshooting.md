# 実行環境と起動の失敗

## 前提の確認の結果

`Test-Prerequisite.ps1`（SKILL.md の「作業前の確認」）が `-OutFile` に書く JSON の読み方。

| 項目 | 中身 | 問題になる時 |
| --- | --- | --- |
| `Ready` | `Problems` が空なら `true` | - |
| `Problems` | 理由と次にすることを、1 件 1 文で | - |
| `Excel` | `Running` と、`Processes`（`Id`、`StartTime`、`WindowTitle`） | Excel が動いている |
| `AccessVBOM` | `Value`（レジストリの値。ポリシーのキー（`HKCU`、`HKLM` の順）を先に読み、利用者の設定のキーは最後。どれにも無ければ `null`）と `Enabled` | `-Vba` を付けていて、値が `1` でない |
| `Wmi` | `Ok` と `Error` | WMI からプロセスを作れない |
| `Workbook` | `Path`、`Exists`、`OpenElsewhere`、`Git`。`-Path` が無ければ `null` | ブックが無い。ブックがほかで開かれている |
| `Workbook.Git` | `Available`（git が使える）、`InRepository`、`Tracked`、`Uncommitted`。分からない項目は `null` | 問題にしない |
| `Error` | 確認そのものが失敗した時の理由 | - |

- `OpenElsewhere` は、ブックの隣に `~$<ブック名>` があるかで決める。Excel は、非表示でも、ブックを開いている間このファイルを置く。Excel が異常終了してファイルが残った時は、開かれていなくても `true` になる。ユーザーに確かめる。
- 非表示の Excel は `WindowTitle` が空になる。題名が空の Excel は、ほかのエージェントやスクリプトが起動したものの可能性がある。
- WMI は、見張りと同じ方法で、すぐ終わる非表示のプロセス（`cmd.exe /c exit`）を 1 つ作って確かめる。
- Excel を閉じた直後は、プロセスが少しの間残り、`Excel.Running` が `true` になることがある。
- 確かめたのは、非表示の Excel でブックを開いている時、無いブック、未追跡のブックと `git add` したブック。`AccessVBOM` が `1` でない時と、WMI が失敗する時は、コマンドを差し替えた試験で確かめた。人が画面に出した Excel、ポリシーで `AccessVBOM` が決まっている端末、ネットワークや OneDrive のフォルダにあるブック、commit 済みのブックでは確かめていない。

## Excel COM の実行環境

- ユーザーが操作できる Windows デスクトップのセッションで実行する。実行環境を変えたときや、そこで COM の成功実績がないときは、作業ブックを開く前に `tests/Test-XlWright.ps1` を流す（SKILL.md の「動作確認」）。
- sandbox の内側で Excel が制限時間内に COM 登録できず、承認された通常の実行では同じ操作が成功した例がある。その差を確認済みの環境では、同じ実行を繰り返さず、対象ブックとコマンドを限定して正式な経路で実行の許可を求める。マクロ・イベント・外部リンク更新を抑止する既定の手順は変えない。
- 明示的な拒否や承認の拒否があれば停止する。実行ポリシーやセキュリティ設定を緩めて続けない。

## Excel COM が起動しないとき

- `New-Object -ComObject Excel.Application` で `0x80080005 (CO_E_SERVER_EXEC_FAILURE)` が出た場合は、対象ブックを開く前の起動失敗として扱う。HRESULT だけでは原因を特定できない。Excel の起動に成功してから `Workbooks.Open` や `VBProject` で失敗した場合は、それぞれの操作を切り分ける。
- 失敗した箇所が不明なときだけ、同じ実行環境で対象ブックを開かずに、`New-ExcelWorkbook` で一時 `.xlsx` を 1 回作ってみる。既に生成単体で失敗を確認している場合は、同じ条件で繰り返さない。
- 起動前後の Excel プロセスを確認する。PID が増えただけでは、COM の生成成功や終了してよいプロセスであることは判断できない。ユーザーに普段の Windows セッションで Excel 単体が起動するか確認してもらう。レジストリ、DCOM、Office の設定は診断の一環として変更しない。
- COM の生成が失敗したままなら、Excel による編集・検証を止め、失敗した操作、HRESULT、対象ブックを開いたか、どこまで確認できたかを報告する。ブック内部を読み取り専用で調べた場合は、その結果と Excel での検証結果を分ける。
- `Could not start the watchdog` が出た場合は、WMI からプロセスを作れない環境で、このスキルは使えない。迂回せずに止まり、人へ伝える。
