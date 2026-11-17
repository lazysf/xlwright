# 作業の前提を、読み取りだけで調べる

- Date: 2026-11-17
- 出典: 一時フォルダの実験用スクリプトを 1 回流した実測。Microsoft 365 Excel 64-bit 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818、Git for Windows。

`SKILL.md` の「作業前の確認」（Excel のプロセス、`AccessVBOM`、Git の状態）と、見張りに要る WMI を、設定もブックも変えずに調べられるかを確かめた。どれも調べられた。WMI からプロセスを作れるかだけは、実際に作らないと分からない。

## 実験の手順

1. `AccessVBOM` を、利用者の設定のキーと、ポリシーのキー（`HKCU` と `HKLM` の `Software\Policies\Microsoft\Office\16.0\Excel\Security`）から読む。
2. WMI で、自分のプロセスを `Get-CimInstance` で引く。`Get-CimClass` で `Win32_Process` に `Create` があるかを見る。見張りと同じ `Win32_Process.Create` で、`cmd.exe /c exit` を非表示で起動する。
3. 新しい `.xlsm` を作る。閉じている時と、非表示の Excel（`Use-ExcelWorkbook`）で開いている時に、次を見る。ブックの隣の `~$<ブック名>` の有無。ブックを共有なし（`FileShare.None`）で開けるか。`Get-Process EXCEL` の結果。
4. Git の外のフォルダと、`git init` したフォルダで、`git rev-parse --is-inside-work-tree`、`git ls-files --error-unmatch`、`git status --porcelain=v1 -- <ブック>` を呼ぶ。ブックが未追跡、追加済み、commit 済み、変更ありの時を比べる。commit は `git write-tree` と `git commit-tree` で作った。

## 結果

### AccessVBOM

- 利用者の設定のキー（`HKCU:\Software\Microsoft\Office\16.0\Excel\Security`）に、値 `1` があった。
- ポリシーのキーは、`HKCU` にも `HKLM` にも無かった。ポリシーで値が決まっている端末での読め方は、確かめていない。

### WMI

| 呼んだもの | 結果 | 時間 |
| --- | --- | --- |
| `Get-CimInstance Win32_Process`（自分の PID） | 自分のプロセスが返った | 136 ms |
| `Get-CimClass Win32_Process` | `Create` がある | 14 ms |
| `Win32_Process.Create`（`cmd.exe /c exit`、非表示） | `ReturnValue` は 0 | 20 ms |

- 問い合わせとクラスの定義は、WMI に届くことしか示さない。プロセスを作れない環境を見つけるには、`Create` を呼ぶしかない。作れない環境での 3 つの結果は、確かめていない。

### ブックが開かれているか

| 状態 | `~$<ブック名>` | 共有なしで開く | `Get-Process EXCEL` |
| --- | --- | --- | --- |
| 閉じている | 無い | 開けた | 無い |
| 非表示の Excel で開いている | ある | `IOException` | PID が 1 つ、`MainWindowTitle` は空 |

- 非表示の Excel でも、開いている間は `~$<ブック名>` ができ、閉じると消えた。
- 非表示の Excel は `MainWindowTitle` が空だった。ウィンドウの題名から、開いているブックは分からない。

### Git

| 状態 | `rev-parse --is-inside-work-tree` | `ls-files --error-unmatch` | `status --porcelain=v1 -- <ブック>` |
| --- | --- | --- | --- |
| Git の外 | 終了コード 128 | - | - |
| 未追跡 | `true` | 終了コード 1 | `?? book.xlsm` |
| 追加済み | `true` | - | `A  book.xlsm` |
| commit 済み | `true` | 終了コード 0 | 空 |
| commit の後に変更 | `true` | 終了コード 0 | `M book.xlsm`（前に空白 1 つ） |

- `-` は呼んでいない組み合わせ。

## Test-Prerequisite.ps1 の形にした理由

使い方の正本は `plugins/xlwright/skills/xlwright/SKILL.md` の「作業前の確認」にある。

- **判定は `XlWright.psm1` の `Get-ExcelPrerequisite` に置き、スクリプトは JSON に書くだけにする。** `AccessVBOM` が 0 の時と WMI が失敗する時は、この端末では起こせない。関数なら、試験がモジュールの中だけでコマンドを差し替えて確かめられる（`watchdog-unavailable` と同じ方法）。
- **WMI は、実際にプロセスを作って確かめる。** 問い合わせは、WMI に届くことしか示さない。作るのは、すぐ終わる非表示の `cmd.exe /c exit` で、見張りと同じ `Start-WmiProcess` を通す。
- **`AccessVBOM` は、ポリシーのキーを先に読む。** `HKCU` のポリシーのキー、`HKLM` のポリシーのキー、利用者の設定のキーの順に読み、最初に値があったものを採る。ポリシーで決まった値を、利用者の設定の値で上書きして報告しないため。2 つのポリシーのキーの順と、ポリシーの値が利用者の設定に勝つことは、Microsoft の文書に書かれていない（2026-11-17 に、トラストセンターとマクロの設定の文書を確かめた）。ポリシーのある端末でも確かめていない。
- **`AccessVBOM` は、`-Vba` を付けた時だけ問題にする。** シートだけの作業で、要らないセキュリティ設定の変更を人に頼まないため。
- **ブックが開かれているかは、`~$<ブック名>` の有無で見る。** 非表示の Excel はウィンドウの題名が空で、題名からは分からない。共有なしで開いてみる方法は、短い間でもブックをロックするので使わない。Excel が異常終了して `~$<ブック名>` が残った時は、開かれていなくても問題として出る。
- **Git の状態は、問題に数えない。** 戻し方（Git か、バックアップか）を決める材料で、作業を止める理由ではない。git が無い端末でも `Ready` になる。
- **問題は、理由と次にすることを 1 文で返す。** Excel を止めない、`AccessVBOM` を自分で変えない、WMI を迂回しない、という `SKILL.md` の規則を、結果を読んだその場で守れるようにする。

`tests/Test-XlWright.ps1` の `prerequisite` が、これを確かめる（2026-11-17）。

- Excel が動いておらず `AccessVBOM` が 1 の時は、`-Vba` を付けても `Ready` になった。
- 非表示の Excel でブックを開いている間は、その PID が `Excel.Processes` に入り、`Workbook.OpenElsewhere` が `true` になり、問題が 2 件出た。ブックのファイルは変わらなかった。
- 無いブックは問題になった。
- `AccessVBOM` を 0 に差し替えると、`-Vba` を付けた時だけ問題になった。`HKLM` のポリシーのキーを 0、利用者の設定のキーを 1 に差し替えると、問題になった。`HKCU` のポリシーのキーを 1、`HKLM` のポリシーのキーを 0 に差し替えると、問題にならなかった。WMI の呼び出しを、例外を投げるものと、終了コード 2 を返すものに差し替えると、どちらも問題になり、理由が `Wmi.Error` に入った。
- git がある端末では、未追跡のブックは `Tracked` が `false`、`git add` したブックは `Tracked` と `Uncommitted` が `true` になり、どちらも `Ready` のままだった。
- ポリシーのキーを読まない写し、`-Vba` を見ない写し、`~$<ブック名>` を見ない写し、Excel を問題にしない写し、WMI の失敗を捨てる写し、無いブックを問題にしない写しは、どれも `prerequisite` で落ちた。
- 非表示の Excel を閉じた直後は、プロセスがまだ残っていて、問題として出ることがあった。試験は、Excel が終わるのを待ってから次を確かめる。

## 未確認

- commit 済みで変更の無いブックの `Uncommitted`。Git のコマンドの結果は実験で見たが、スクリプトを通しては確かめていない。試験で commit を作ると、端末の Git の設定（名義、署名）に左右される。
- ポリシーで `AccessVBOM` が決まっている端末。ポリシーの値が利用者の設定に勝つか、`HKCU` と `HKLM` のポリシーのどちらが勝つか。
- WMI からプロセスを作れない環境。
- 人が画面に出した Excel でブックを開いている時の `~$<ブック名>` と `MainWindowTitle`。
- ネットワークドライブや OneDrive の同期フォルダにあるブック。
- `.gitignore` で無視されているブック、Git LFS、worktree。
- 32-bit の Excel と、`16.0` 以外のキーを使う版。
