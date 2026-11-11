---
name: xlwright
description: Excel ブックを Windows 版 Excel と PowerShell COM で確認・編集する。シートの値・数式・書式、VBA と UserForm の作成・編集と入出力、カスタムリボンの埋め込みとボタンの動作確認、VBA をロックした配布ブックの出力に使う。
---

# xlwright

PowerShell から `Excel.Application` を COM で操作してブックを編集する。リボンだけは COM で定義できないため、閉じたブックの zip を直接書き換える。

このスキルのフォルダは `${CLAUDE_SKILL_DIR}` で、スクリプトはその `scripts/` にある。`references/` のコマンド例にある `<skill>` は、このフォルダの絶対パスに置き換えて実行する。Git Bash からは `powershell.exe -NoProfile -File "<script>"` で実行し、引数のパスは `/` 区切りのまま渡してよい。

## 利用の前提

- Windows 版 Excel（デスクトップ版）、Excel の COM、Windows PowerShell 5.1 を使う。
- WMI（`Win32_Process.Create`）からプロセスを作れること。Excel を起動するたびに、上限で Excel を止める見張りのプロセスを WMI で立てる。立てられない環境では、このスキルは使えない。`Could not start the watchdog` のエラーが出たら、迂回せずに止まり、人へ伝える。
- エージェントの作業中、人はその端末で Excel を使わない。
- 利用する repo に、ブックの置き場や VBA ファイルの規則（`AGENTS.md`、`CLAUDE.md` など）があれば、そちらを優先する。無ければ、このスキルの既定に従う。

## 作業前の確認

1. **Excel のプロセス**: `Get-Process EXCEL`。動いているなら、ユーザーが開いているブックがある可能性がある。勝手に終了させず、閉じてよいか聞く。
2. **VBA を触るなら `AccessVBOM`**: `HKCU:\Software\Microsoft\Office\16.0\Excel\Security` の `AccessVBOM` が `1` でないと、`VBProject` へのアクセスが失敗する。これはセキュリティ設定なので、自分で変更せず、ユーザーに有効化を頼む。ユーザーに頼まれた時も、確認を取った後も、エージェントは AccessVBOM などのトラストセンターの設定を変えない。手順を渡して、ユーザーに変えてもらう。
3. **バックアップ**: 編集前のブックを scratchpad などの一時領域へコピーする。段階ごと（COM 編集後、customUI 埋め込み前など）にも取ると、どこで壊れたか切り分けられる。
4. **Git 管理**: ブックが Git で追跡されているか、未 commit の変更があるかを確かめる。未 commit の編集や追跡されないブックは、Git では戻せないのでバックアップから戻す。

初めての端末や、Excel が更新された後は、[実行環境と起動の失敗](references/troubleshooting.md)に従い、先に「動作確認」の試験を流す。

## マクロが動くことへの注意

Excel は既定で、VBA とブックイベントを無効にし、外部ブックリンクを更新せずに開く。この設定は Excel 4.0 マクロやデータ接続の更新全般を抑止するものではない。

- VBA の動作検証に限り `-EnableMacros` を付ける。実行前に `onLoad` などのコールバックも確認する。`Workbook_Open` などのイベントを検証する時は `-EnableEvents` も付ける。
- `$xl.Run` やボタンの押下で動かしてよいのは、この作業で自分が書いた、外部とやり取りしないマクロだけ。既存のマクロやボタンは、ユーザーの確認なしに動かさない。サーバーやファイルを読み書きするマクロは、ブックを保存せずに閉じても、その変更は戻らない。
- sandbox を、マクロに対する防御として数えない。COM で起動した Excel が sandbox の制限の内側で動くとは限らない。

## 手段の分担と順序

| 対象 | 手段 | 理由 |
| --- | --- | --- |
| シート（値、数式、書式、シート構成） | COM（`Invoke-ExcelScript.ps1`） | Excel が再計算し、保存形式も Excel が保つ |
| VBA モジュール、UserForm | COM（`Export-Vba.ps1` / `Import-Vba.ps1`） | `vbaProject.bin` の生成を Excel に任せられる |
| リボン（customUI） | zip を直接編集（`Add-CustomUI.ps1`） | COM にリボンを定義する API が無い |

順序は **COM の編集 → ブックを閉じる → customUI の埋め込み → COM で開き直して検証**。openpyxl など Excel 以外のツールで保存しない（リボンの関連付けが消える。[リボン](references/ribbon.md)）。

## 対象と作業手順

ブックごとにフォルダを分け、ブックの隣にソースを置く配置を既定とする。各スクリプトの `-Path` には、そのフォルダのブックを渡す。対象が特定できない場合は、作業ブックの場所を確認してから開く。

```text
<ブックのフォルダ>/
├── <ブック>.xlsm
├── vba/
│   ├── standard/
│   ├── classes/
│   ├── documents/
│   └── forms/
└── ribbon/
    └── customUI14.xml
```

| 作業 | 読むもの |
| --- | --- |
| シートの確認・編集 | 下の「COM 編集」。確認だけなら `-Save` を付けない |
| VBA の変更（`vba/` の正本を編集して取り込む） | [VBA の編集ルールと入出力](references/vba.md) |
| UserForm の作成・変更 | [UserForm の作成と編集](references/userform.md) |
| リボンの変更とボタンの動作確認 | [リボン](references/ribbon.md) |
| VBA をロックした配布ブックの出力 | [配布ブックの出力](references/publish.md) |
| 複数のブックを順に扱う、別ブックからシートをコピーする | [COM 編集の詳細](references/com.md) |
| Excel が起動しない、実行が拒否される | [実行環境と起動の失敗](references/troubleshooting.md) |

編集後はブックを開き直し、「検証」で変更した箇所を確認する。

### 新しいブックを作る時の流れ

VBA とリボンを持つブックを一から作る時は、次の順に進める。各手順の詳細は上の表の reference にある。

1. `Invoke-ExcelScript.ps1 -Create -Save` でブックを作り、シートの値・数式・書式を入れる。
2. `vba/` に VBA のソースを書き、`Import-Vba.ps1` で取り込む。
3. ブックを閉じた状態で、`ribbon/customUI14.xml` を `Add-CustomUI.ps1` で埋め込む。
4. `Invoke-ExcelScript.ps1` で開き直し、シートと VBA を読み戻す。リボンは `Invoke-RibbonButton.ps1` で押して確かめる。
5. 配布する場合は、最後に `Publish-Workbook.ps1` で出力し、ロックを読み戻す。

## COM 編集

Excel を `New-Object -ComObject` で直接起動しない。起動、ブックを開く、後始末は `XlWright.psm1` の `Use-ExcelWorkbook` に任せる。使い捨ての操作は、ScriptBlock の中身だけをファイルに書き、`Invoke-ExcelScript.ps1` に渡す。ファイルは scratchpad などの一時領域に置く。

```powershell
# read.ps1（UTF-8。日本語を書いてよい）
param($xl, $wb)
$ws = $wb.Worksheets.Item('集計表')
@{ A1 = $ws.Range('A1').Value2; B1 = $ws.Range('B1').Formula; Rows = $ws.Range('A1:B2').Value2 }
```

```powershell
powershell.exe -NoProfile -File "${CLAUDE_SKILL_DIR}/scripts/Invoke-ExcelScript.ps1" -Path <workbook> -Script <read.ps1> -OutFile <result.json>
```

- スクリプトのファイルは UTF-8 として読む（BOM の有無を問わない）。日本語のシート名や文字列を直接書いてよい。
- 結果は UTF-8 の JSON で `-OutFile` に書く。`Ok`、`Result`（ScriptBlock が出力した値の配列）、`Error`、`Warnings` を持つ。失敗した時は終了コードが 1。`powershell.exe` の標準出力は Git Bash から見ると日本語が化けるので、結果は JSON で読む。
- 返す値は文字列・数値・真偽値・hashtable・配列にする。COM オブジェクトは `"[COM object]"` に置き換わる。複数セルの `Value2`（2 次元配列）は行の配列になる。COM メソッドの戻り値（`Worksheets.Add()` など）も結果に混ざるので、要らないものは `[void]` で捨てる。
- `-Create` を付けると、先に新しいブックを作る（既にあれば失敗する）。拡張子が `.xlsm` ならマクロ有効ブック、`.xlsx` なら通常のブックになる。
- `-Save` は、ScriptBlock が例外なしで終わった時だけ保存する。例外が出たら、保存せずに閉じる。
- `-Visible` はウィンドウを表示する。ユーザーの画面に Excel が出るので、先に一言伝える。
- `-TimeoutSeconds`（既定 300）を過ぎると、見張りのプロセスがこの呼び出しで起動した Excel だけを止め、`Timed out after ...` のエラーになる。保存していない変更は失われる。見張りは呼び出し元のシェルが timeout などで殺されても残り、上限で Excel を止める。長いマクロを動かす時は、上限を明示して延ばす。
- ブックを閉じて `Quit` した後、Excel が 10 秒以内に終わらなければ、この呼び出しが起動した Excel（PID と起動時刻で照合）を止める。PowerShell が COM の参照を手放さず、`Quit` の後も Excel が数分残る実測があるため。
- 警告で PID が出た時（照合できない、止められない）は、人へ伝える。自分で止める時も、その PID だけを `Stop-Process -Id` で止める。名前（`-Name EXCEL`）では止めない。ユーザーの Excel まで止まる。

### PowerShell + COM の癖

- **セルへの代入は `Range('B2')` 形式で、型を明示する。** `$ws.Cells.Item(r, c).Value2 = 3` は `System.Int32` を `System.String` にキャストできないという `InvalidCastException` で落ちることがある。`$ws.Range("B$r").Value2 = [double]$v` なら通る。
- **`-File` で直接実行するスクリプトは ASCII だけで書く。** `powershell.exe`（5.1）は BOM なしの `.ps1` を UTF-8 として読まない。`Invoke-ExcelScript.ps1` に渡すファイルには、この制限はない。

## 検証

編集した後は、`-Save` なしでブックを開き直して読み戻す。書いたつもりの値ではなく、Excel が保存した結果を見る。動かしてよいマクロの範囲は「マクロが動くことへの注意」のとおり。

- **シート**: `Value2`、`Formula`、`Text`（表示形式が適用された文字列）を読む。数式は Excel が再計算した値で確かめる。
- **VBA**: 変更したコンポーネントの名前・種別・コードを読み戻す。実行してよいマクロは `$xl.Run('MacroName')` で戻り値を確かめる。ThisWorkbook に置いた関数は、`$xl.Run` で呼ぶと（ブック名で修飾しても）空が返る。`[System.__ComObject].InvokeMember('FuncName', [Reflection.BindingFlags]::InvokeMethod, $null, $wb, $null)` で Workbook のメンバーとして呼ぶと値が返る。
- **リボンとボタン**: `Invoke-RibbonButton.ps1` で実際に押し、出たダイアログの文言を確かめる（[リボン](references/ribbon.md)）。

## 動作確認

スクリプトが手元の Excel で動くかは、`tests/Test-XlWright.ps1` で確かめる。各項目は一時フォルダに新しいブックを作って、期待値と照合する。一時フォルダの外は変えない。Excel が動いている時と `AccessVBOM` が `1` でない時は、何もせずに終了コード 2 で止まる。

```powershell
powershell.exe -NoProfile -File "${CLAUDE_SKILL_DIR}/tests/Test-XlWright.ps1" -OutFile <result.json>
```

- 既定の項目は Excel を表示しない。セルと VBA、`Invoke-ExcelScript.ps1`、上限、呼び出し元が殺された時、不正な入力の拒否（VBA のソース、リボンの XML、配布の失敗）、見張りが立てられない時、UserForm を確かめる。5 分ほどかかる。
- `-Full` を付けると、リボンの押下（同名の組み合わせ、閉じられないダイアログを含む）と配布ブックの出力も確かめる。Excel と VBE が画面に出るので、先に人へ伝え、終わるまで操作しないよう頼む。
- `-Case timeout,userform` のように、名前を指定した項目だけを流せる。
- 終了コードは、0 がすべて合格、1 が不合格あり、2 が前提を満たさない・Excel が残って止まった。不合格の時は一時フォルダを残し、場所を出力する。
