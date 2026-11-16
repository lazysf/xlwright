# VBA の編集ルール

コマンド例の `<skill>` は、SKILL.md に書いたこのスキルのフォルダの絶対パスに置き換える。

## 保存形式とモジュール属性

- `.bas`、`.cls`、`.frm` の正本は、UTF-8（BOM なし）・LF で保存する。利用する repo に別の規則があれば、そちらに従う。VBE との文字コード・改行の変換は入出力スクリプトに任せる。
- `.bas` は先頭行を `Attribute VB_Name = "<モジュール名>"` にし、その前に空行やクラスヘッダー（`VERSION 1.0 CLASS` など）を置かない。
- ファイル名（拡張子を除く）は `VB_Name` と一致させる。`Import-Vba.ps1` は一致しないファイルを拒否する。
- 書き出したファイルを編集する時は、`Attribute VB_Name` などのモジュール属性を保持する。新しいクラスの `.cls` には、VBE が取り込めるクラスヘッダーと属性を含める。
- 文書モジュールの属性は書き出したファイルに保持する。シートとブックへの反映では、既存モジュールのコード部分を差し替える。

## 識別子

- モジュール名（`Attribute VB_Name`）は31文字以内にする。プロシージャや変数などの一般識別子は255文字以内にする。
- 識別子は文字で始める。空白、ピリオド、`!`、`@`、`&`、`$`、`#` を含めない。予約語と同一スコープでの名前の衝突を避ける。
- `Option Explicit` はモジュールの宣言部に置く。
- モジュール名は役割が分かる名前にする（例: リボンの callback を置くなら `RibbonCallbacks`）。`Ribbon` のように、変数名（`IRibbonUI` を入れる変数など）やプロシージャ名とぶつかりやすい名前は避ける。

## 変更前の確認

- 公開プロシージャ、マクロの入口、シート名やモジュール名を変更する前に、他のモジュール、シートの数式、リボン XML からの参照を確認する。
- ブックのフォルダの `vba/` にある正本とブック内のコードの同期が不明な場合は、ブックから別の作業先に書き出して差分を確認する。どの変更を残すか判断できない場合は、反映前に確認する。

## 書き出しと取り込み

VBA をテキストで直す時や、ソースを Git で管理する時に使う。標準・クラス・文書モジュールと UserForm を扱う。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Export-Vba.ps1" -Path <workbook> -OutDir <dir>
powershell.exe -NoProfile -File "<skill>/scripts/Import-Vba.ps1" -Path <workbook> -SourceDir <dir>
```

- **Export**: 標準モジュールは `standard/<名前>.bas`、クラスは `classes/<名前>.cls`、文書モジュールは `documents/<名前>.cls`、UserForm は `forms/` に `.frm` と付随する `.frx` で書き出す。コードの無い文書モジュールは省く。今回書き出さなかったファイルが出力先に残っていれば警告する。Excel が一時出力した CP932・CRLF のファイルを UTF-8（BOM なし）・LF に変換する。
- **Import**: `-SourceDir` の `standard/`、`classes/`、`documents/`、`forms/` から種別を決める。ブックを開く前に、すべてのファイルが UTF-8（BOM なし）・LF で、CP932 に文字を失わず変換できることを確認する。VBE へは CP932・CRLF の一時ファイルを渡す（LF の `.cls` を VBE へ直接渡すと、標準モジュールとして認識された実測がある）。同名・同種別の標準モジュール、クラス、UserForm は削除してから取り込む。文書モジュールは既存のコードを差し替え、対象ブックに同名の文書モジュールが無い場合は止まる。取り込み元に含まれないコンポーネントはブックに残る。最後に Excel で保存する。
- 文書モジュールの差し替えは、先頭のヘッダーと `Attribute` 行だけを除く。プロシージャの中の `Attribute` 行（`VB_Description` など）は扱わない。
- 変更したコンポーネントだけを反映する場合は、対象ファイルを種別ごとの配置を保って別の作業先へコピーし、そのルートを `-SourceDir` に渡す。UserForm は付随する `.frx` もコピーする。
- シートとブックのモジュールを直接編集する場合は、`Worksheet.CodeName` または `Workbook.CodeName` で既存の `VBComponent` を取得し、その `CodeModule` に書く。
- 既存のブックから書き出したコードには、接続先、ID、認証情報が入っていることがある。commit する前に中身を確かめる。

## コンパイルの検査

取り込んだ後に、VBA プロジェクトをコンパイルして、構文エラーや宣言の無い変数を実行の前に見つける。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Test-VbaCompile.ps1" -Path <workbook> -OutFile <result.json>
```

- マクロを無効にしてブックを開き、VBE の「VBAProject のコンパイル」を実行して、保存せずに閉じる。ブックは変わらない。
- 結果は UTF-8 の JSON で `-OutFile` に書く。`Compiled` は、通れば `true`、コンパイルエラーなら `false`、検査そのものが失敗した時は `null`。終了コードは、通った時だけ 0。
- コンパイルエラーの時は、`CompileError` に `Message`（「コンパイル エラー: 変数が定義されていません。」など、VBE のダイアログの文言）、`Module`、`Line`、`Column` が入る。`Line` は VBE のコードウィンドウの行で、ファイルの先頭の `Attribute` 行やクラスヘッダーを数えない（`.bas` なら、ファイルの行番号から 1 を引いた値）。
- 報告するエラーは 1 つだけ。直したら、通るまで流し直す。
- `Module`、`Line`、`Column` は、VBE から場所を読めなかった時は `null` になる。宣言の無い変数と構文エラーでは読めた。
- コンパイルエラーがある時だけ、VBE のウィンドウとダイアログが画面に短い間出る。スクリプトが閉じて隠す。
- 検査そのものの失敗は `Error` に入る。VBA プロジェクトがロックされている時、`-DialogTimeoutSeconds`（既定 5）の間にコンパイルの済んだ状態にもならずダイアログも出なかった時など。後者は合格として扱わない。大きなプロジェクトでこの失敗が出たら、秒数を延ばす。
- VBA の無いブック（`.xlsx` を含む）は `Compiled` が `true` になる。
- 確かめたのは、標準モジュールにエラーが 1 つある場合。エラーが複数ある時、文書モジュールや UserForm にエラーがある時、VBA プロジェクトがロックされている時、日本語以外の表示言語の Excel では確かめていない。

## テストの実行

書いた VBA の動きを、VBA で書いたテストで確かめる。テストは、ブックの標準モジュールに置く `Sub` で、次をすべて満たすもの。

- 名前が `Test_` で始まる。
- 引数が無い。
- `Private` でない。

テストは、期待と違う時に `Err.Raise` でエラーを上げる。最後まで進めば合格になる。

```vb
Attribute VB_Name = "Test_Util"
Option Explicit

Public Sub Test_AddNums()
    Dim actual As Long
    actual = Util.AddNums(2, 3)
    If actual <> 5 Then Err.Raise vbObjectError + 1, , "AddNums(2, 3): expected 5, actual " & actual
End Sub
```

`vba/standard/` に置いて `Import-Vba.ps1` で取り込み、`Test-VbaCompile.ps1` でコンパイルを確かめてから実行する。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Invoke-VbaTest.ps1" -Path <workbook> -OutFile <result.json>
```

- マクロを有効、イベントを無効にしてブックを開き、テストを 1 つずつ実行して、保存せずに閉じる。ブックのファイルは変わらない。テストがほかのファイルやサーバーに書いた変更は戻らない。動かしてよいマクロの範囲は、SKILL.md の「マクロが動くことへの注意」のとおり。
- テストは、モジュール名の順、モジュールの中ではコードの順に実行する。`-Name Test_Util.Test_AddNums,Test_Util.Test_Other` のように `<モジュール>.<Sub>` を渡すと、そのテストだけを実行する。
- 結果は UTF-8 の JSON で `-OutFile` に書く。終了コードは、すべてのテストが合格した時だけ 0。

  ```json
  {
    "Ok": false,
    "Summary": { "Total": 2, "Passed": 1, "Failed": 1, "NotRun": 0 },
    "Tests": [
      { "Name": "Test_Util.Test_AddNums", "Result": "passed", "ElapsedMs": 15, "Message": null },
      { "Name": "Test_Util.Test_Other", "Result": "failed", "ElapsedMs": 90, "Message": "実行時エラー '-2147221503 (80040001)': expected 5, actual 6" }
    ],
    "Error": null,
    "Warnings": []
  }
  ```

- `Result` は、`passed`、`failed`、`not_run` のどれか。`failed` の `Message` は、実行時エラーのダイアログの文言で、エラーの番号と `Err.Raise` の説明が入る。説明の改行は空白になる。`Err.Raise` の `Source` は返らない。
- テストが自分で上げたエラーと、0 での除算のような意図しない実行時エラーは、区別せずに `failed` になる。
- 失敗したテストごとに、実行時エラーのダイアログが画面に短い間出る。スクリプトが閉じるので、人は操作しない。
- 実行時エラーの後は、モジュール変数が初期値に戻る。テストは、前のテストが残した状態に頼らない形で書く。
- コンパイルエラーのあるテストは `failed` になり、`Message` にコンパイルエラーの文言が入る。VBE のウィンドウが短い間出る。残りのテストは続けて実行する。
- OK だけの MsgBox は閉じて先へ進み、文言が `Warnings` に入る。はい/いいえ の MsgBox のような閉じられないダイアログでは、そのテストを `failed` にして Excel を止める。残りのテストは `not_run` になり、`Error` に理由が入る。`-TimeoutSeconds`（既定 300）を過ぎた時も同じ。
- テストが 1 つも無い時、`-Name` のテストが無い時、VBA プロジェクトがロックされている時は、`Error` に理由が入り、終了コードは 1 になる。
- テストのモジュールは、ブックの中に残る。配布ブックに入れたくない時は、出力の前にブックから消す。
- 確かめたのは、標準モジュールのテストを日本語の表示の Excel で実行した場合。VBA プロジェクトがロックされたブック、VBE のオプションで「エラー発生時に中断」を選んでいる時は確かめていない。
