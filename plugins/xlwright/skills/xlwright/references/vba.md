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
- ブックのフォルダの `vba/` にある正本とブック内のコードの同期が不明な場合は、ブックを書き出して正本と比べる（「正本との照合」）。どの変更を残すか判断できない場合は、反映前に確認する。

## 書き出しと取り込み

VBA をテキストで直す時や、ソースを Git で管理する時に使う。標準・クラス・文書モジュールと UserForm を扱う。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Export-Vba.ps1" -Path <workbook> -OutDir <dir>
powershell.exe -NoProfile -File "<skill>/scripts/Import-Vba.ps1" -Path <workbook> -SourceDir <dir>
```

- **Export**: 標準モジュールは `standard/<名前>.bas`、クラスは `classes/<名前>.cls`、文書モジュールは `documents/<名前>.cls`、UserForm は `forms/` に `.frm` と付随する `.frx` で書き出す。コードの無い文書モジュールは省く。今回書き出さなかったファイルが出力先に残っていれば警告する。Excel が一時出力した CP932・CRLF のファイルを UTF-8（BOM なし）・LF に変換する。参照設定（VBE の「ツール」→「参照設定」）は、出力先の `references.json` に書く（下の「参照設定」）。
- **Import**: `-SourceDir` の `standard/`、`classes/`、`documents/`、`forms/` から種別を決める。ブックを開く前に、すべてのファイルが UTF-8（BOM なし）・LF で、CP932 に文字を失わず変換できることを確認する。VBE へは CP932・CRLF の一時ファイルを渡す（LF の `.cls` を VBE へ直接渡すと、標準モジュールとして認識された実測がある）。同名・同種別の標準モジュール、クラス、UserForm は削除してから取り込む。文書モジュールは既存のコードを差し替え、対象ブックに同名の文書モジュールが無い場合は止まる。取り込み元に含まれないコンポーネントはブックに残る。最後に Excel で保存する。途中で失敗した時は、保存せずに閉じる。
- **参照設定**: 書き出しのファイルには参照設定が出ず、取り込みでも足されない。参照設定の要るコード（`Dim d As Scripting.Dictionary` など）は、正本から作り直したブックではコンパイルが通らない。そこで、`references.json` に参照設定を残し、取り込みで照らす。
  - `references.json` は `{"References": [{"Name", "Description", "Guid", "Major", "Minor"}, ...]}` の形で、`VBProject.References` の順に並ぶ。外せない VBA と Excel は書かない。新しいブックにも最初からある `stdole` と `Office` は書く。端末で変わる `FullPath` は書かない。ほかのブックへの参照は `Guid` を持たないので書かず、参照先の `FullPath` を `Not recorded in references.json: <FullPath>` と警告する。アドイン（`.xlam`）への参照も同じに扱うが、確かめたのは `.xlsm` への参照だけ。
  - 取り込みは、ブックを開く前に `references.json` の形を確かめ、各項目に `Name` と `Guid` が無ければ止まる。取り込んだ後に、ブックに無い参照設定（`Guid` で照らし、版は比べない）を `Reference not in the workbook: <Name> <Guid>` と警告する。参照設定は足しも外しもしない。`references.json` が無い正本は照らさない。
  - 警告が出たら、ユーザーに参照設定の名前を示して確認を取ってから足す。足すのは、`Invoke-ExcelScript.ps1 -Save` に渡す ScriptBlock で `[void]$wb.VBProject.References.AddFromGuid('<Guid>', <Major>, <Minor>)` を呼ぶ形でよい。Scripting Runtime では、確認のダイアログを出さずに足せた。参照先が端末に無い時、版が違う時の振る舞いは確かめていない。足した後に `Test-VbaCompile.ps1` を流す。
  - 新しく書くコードで参照設定を避けたい時は、`CreateObject("Scripting.Dictionary")` と `As Object` の遅延バインディングで書けば、参照設定なしでコンパイルが通る。
- 文書モジュールの差し替えは、先頭のヘッダーと `Attribute` 行だけを除く。プロシージャの中の `Attribute` 行（`VB_Description` など）は扱わない。
- 変更したコンポーネントだけを反映する場合は、対象ファイルを種別ごとの配置を保って別の作業先へコピーし、そのルートを `-SourceDir` に渡す。UserForm は付随する `.frx` もコピーする。
- シートとブックのモジュールを直接編集する場合は、`Worksheet.CodeName` または `Workbook.CodeName` で既存の `VBComponent` を取得し、その `CodeModule` に書く。
- 既存のブックから書き出したコードには、接続先、ID、認証情報が入っていることがある。commit する前に中身を確かめる。

### 正本から消したモジュール

取り込みは、正本に無いコンポーネントを消さない。正本からモジュールを消した時は、ブックの同じコンポーネントを名前で指定して消す。消す前に「正本との照合」でブックにだけあるものを調べ、名前をユーザーに示して確認を取る（SKILL.md の「対象と作業手順」）。確認した名前だけを、`Invoke-ExcelScript.ps1 -Save` に渡す ScriptBlock で消す。

```powershell
param($xl, $wb)
$components = $wb.VBProject.VBComponents
foreach ($name in 'OldModule', 'Sheet1') {
    $c = $components.Item($name)
    if ($c.Type -eq 100) {
        # A document module cannot be removed; with no code, it is not exported.
        $cm = $c.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        'cleared: ' + $name
    }
    else {
        $components.Remove($c)
        'removed: ' + $name
    }
}
```

- 標準モジュール、クラス、UserForm は `Remove` で消える。確認のダイアログは出ない。
- 文書モジュール（ThisWorkbook、シート）は、`Remove` を呼ぶと `ArgumentException` になる。コードを空にする（シートは消さない）。コードの無い文書モジュールは書き出されないので、照合の差から外れる。
- ブックに無い名前を渡すと `Item` が失敗し（「インデックスが有効範囲にありません。」）、`-Save` は保存しない。その前の名前で消した分も、ブックに残らない。
- 消したモジュールを呼ぶコードが残っていても、`Remove` と保存は通る。後に `Test-VbaCompile.ps1` でコンパイルを確かめる。
- 確かめたのは、標準モジュール、クラス、ラベル 1 つの UserForm、シートのモジュール。ほかの種別のコンポーネント（ActiveX デザイナーなど）、ロックされた VBA プロジェクトでは確かめていない。この例の ScriptBlock そのものは試験に無い。

## 正本との照合

正本（`vba/`）とブックの中のコードが同じかを、取り込む前や、人がブックを直した後に確かめる。ブックを新しい一時フォルダへ書き出して、正本と `git diff` で比べる。ブックも正本も変えない。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Export-Vba.ps1" -Path <workbook> -OutDir <new temp dir>
git diff --no-index --ignore-blank-lines <vba> <new temp dir>
```

- 書き出し先は、毎回新しいフォルダにする。前の書き出しが残っていると、ブックから消えたコンポーネントが差に出ない。
- 正本にだけあるファイルは削除（`deleted file`）、ブックにだけあるコンポーネントは追加（`new file`）として出る。
- 差が無ければ何も出ず、終了コードは 0 になる。UserForm があると、`.frx` が書き出すたびに違うバイト列になるので、`Binary files ... differ` が出て終了コードは 1 になる。この行は無視する。`.frx` のほかに差が無ければ一致とみなす。コントロールの配置やフォントの差は、`.frm` の先頭のプロパティに出る分しか分からない。
- `--ignore-blank-lines` を付けるのは、VBE が `.frm` を取り込むたびに `Attribute` の行の後へ空行を 1 つ足すため。このオプションは、コードの途中の空行の差も隠す。
- VBE は取り込む時に、キーワードの大文字と小文字、演算子の前後の空白、行末の空白を書き換え、ファイルの末尾に改行を足す。VBE の形で書いていない正本は差になる。一度書き出して、正本を書き出しの形にすると、この差は消える。
- 大文字と小文字だけの差は、VBE が識別子の大文字と小文字をプロジェクトの全体でそろえた結果のことが多い。あるモジュールを取り込むと、触っていないモジュールの識別子が変わることがある。文字列の中身を変えた差と見分けがつかないので、行を読んで判断する。
- モジュール名の大文字と小文字だけが違うと（`Util.bas` と `util.bas`）、同じモジュールが削除と追加として出る。
- 書き出しには `references.json` も入るので、参照設定の違いもこのファイルの差として出る。
- `vba/` に、VBA と `references.json` のほかのファイルがあると、それも削除として出る。
- 確かめたのは、標準モジュール、クラス、文書モジュール、コントロールとコードを持つ UserForm（git 2.56）。この手順は試験に無い。

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
