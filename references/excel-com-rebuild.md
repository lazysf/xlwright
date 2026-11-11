# excel-com スキルの復元

この Markdown 1 本から、`.agents/skills/excel-com/` の 26 ファイルを復元できます。各ファイルの相対パス、SHA-256、原文を収録しています。Excel ブックやリポジトリの設定は収録対象外です。

## AI に復元を依頼する

この Markdown ファイルを添付し、次のように依頼します。

> 添付した Markdown の「PowerShell で復元する」に従い、指定した出力先に `excel-com` フォルダを作成してください。「原文」の 26 ファイルを相対パスどおりに書き出し、各ファイルの SHA-256 が記載値と一致することを確認してください。内容の要約・修正・補完や、復元したスクリプトの実行は不要です。

## PowerShell で復元する

出力先の親フォルダを PowerShell で開きます。リポジトリ内で使う場合は `.agents/skills/` を親フォルダにします。下のコードの `$bundlePath` をダウンロードした Markdown のパスに変更し、コードブロックを実行します。既に `excel-com` フォルダがある場合は、比較用に別の親フォルダを選んでください。

復元コードは原文を UTF-8（BOM なし）・LF で書き出します。全ファイルの SHA-256 を書き込み前に照合し、既存の出力先があれば停止します。

~~~powershell
$bundlePath = '.\excel-com-rebuild.md'
$ErrorActionPreference = 'Stop'
$utf8 = [System.Text.UTF8Encoding]::new($false, $true)
$bundle = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $bundlePath).ProviderPath, $utf8)
$bundle = $bundle.Replace([string]([char]13) + [char]10, [string][char]10)

$record = '(?m)^<!-- excel-com-file (?<path>(?:[A-Za-z0-9_-]+/)*[A-Za-z0-9_-]+\.(?:md|yaml|ps1|psm1)) (?<hash>[a-f0-9]{64}) -->\n~~~~~[a-z]+\n(?<body>.*?)^~~~~~$'
$entries = [regex]::Matches($bundle, $record, [Text.RegularExpressions.RegexOptions]::Singleline)
if ($entries.Count -ne 26) {
    throw "復元対象は 26 ファイルです。検出数: $($entries.Count)"
}

$files = @{}
$sha256 = [Security.Cryptography.SHA256]::Create()
try {
    foreach ($entry in $entries) {
        $relativePath = $entry.Groups['path'].Value
        if ($files.ContainsKey($relativePath)) {
            throw "相対パスが重複しています: $relativePath"
        }
        $bytes = $utf8.GetBytes($entry.Groups['body'].Value)
        $hash = [BitConverter]::ToString($sha256.ComputeHash($bytes)).Replace('-', '')
        if ($hash -ne $entry.Groups['hash'].Value) {
            throw "SHA-256 が一致しません: $relativePath"
        }
        $files.Add($relativePath, $bytes)
    }
}
finally {
    $sha256.Dispose()
}

$output = Join-Path (Get-Location).ProviderPath 'excel-com'
if (Test-Path -LiteralPath $output) {
    throw "出力先が既に存在します: $output"
}
New-Item -ItemType Directory -Path $output -ErrorAction Stop | Out-Null
foreach ($relativePath in ($files.Keys | Sort-Object)) {
    $target = Join-Path $output $relativePath
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target)) | Out-Null
    [IO.File]::WriteAllBytes($target, $files[$relativePath])
}
Write-Output "$($files.Count) ファイルを復元しました。"
~~~

## 原文

### evals/export-edit-import-vba/case.yaml

<!-- excel-com-file evals/export-edit-import-vba/case.yaml 18e39a8f1ffa507e6aca1352a6383376716bfb1217aa4c2b8fa0d6f53f122cd8 -->
~~~~~yaml
schema_version: "1.1"
name: export-edit-import-vba
expected_outcome: |
  calc.xlsm の標準モジュール Util の AddNums は a+b を返す。クラスモジュール CalcOptions、先頭シートのモジュール、ThisWorkbook の編集が保存後も反映されている。
  src/excel/calc/vba/standard/Util.bas が AGENTS.md のルール（UTF-8、BOM なし、先頭に空行なし、LF、クラスヘッダーなし、VB_Name = "Util"）を満たし、日本語コメントが入っている。
  src/excel/calc/vba/classes/ に CalcOptions.cls、documents/ に先頭シートと ThisWorkbook の .cls がある。
  取り込み直した後の標準モジュールは Util、クラスモジュールは CalcOptions のみ。先頭シートと ThisWorkbook は元の文書モジュールのまま。日本語コメントは化けていない。
  開き直したブックの先頭シートで、A1 の値が 2、B1 の数式が =A1*3、B1 の計算結果が 6。
  作業後に Excel のプロセスが残っていない。
~~~~~

### evals/export-edit-import-vba/graders/01-bas-follows-repo-rules.md

<!-- excel-com-file evals/export-edit-import-vba/graders/01-bas-follows-repo-rules.md 754b1a6bc841c5c311efb7b7a90b3d82cb1de48c0868eac62723d72afe09235f -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

src/excel/calc/vba/standard/Util.bas は BOM なし、先頭行が Attribute VB_Name = "Util"、改行が LF のみ、UTF-8 として読めて「2 つの整数を足す」というコメントを含み、VERSION 1.0 CLASS などのクラスヘッダーを含まない
~~~~~

### evals/export-edit-import-vba/graders/02-no-duplicate-module.md

<!-- excel-com-file evals/export-edit-import-vba/graders/02-no-duplicate-module.md 74b75dd8c097869183982bc5ea0bbad1cbb2df04c657658322b7062d1f7e3eae -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

取り込み直した後の calc.xlsm の標準モジュールは Util のみ、クラスモジュールは CalcOptions のみ。CalcOptions の Precision は String で、Util1 や ThisWorkbook1 のような重複コンポーネントが無い。
~~~~~

### evals/export-edit-import-vba/graders/03-japanese-comment-intact.md

<!-- excel-com-file evals/export-edit-import-vba/graders/03-japanese-comment-intact.md 94f88b435c6ee95068e29f30640a3b287ac4bb3085e413d8a888ce1c14bbdb7f -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

取り込み直した後の calc.xlsm の Util に「2 つの整数を足す」というコメントが化けずに入っている（CodeModule から読み戻して一致する）
~~~~~

### evals/export-edit-import-vba/graders/04-function-still-works.md

<!-- excel-com-file evals/export-edit-import-vba/graders/04-function-still-works.md 4f5e96fab133ee3e0fb4c376f3b866935036d263ca0225e5c653a2c0c1494860 -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

取り込み直した後に AddNums を実行し、a+b の値が返ることを確かめている
~~~~~

### evals/export-edit-import-vba/graders/05-no-excel-left.md

<!-- excel-com-file evals/export-edit-import-vba/graders/05-no-excel-left.md 5ccff1bd657c0a71f5075b4e95d2e44f65cf4df002f5d68d26ef2ef9e46f3b5d -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

作業の終了時点で、実行中に起動した EXCEL.EXE が残っていない
~~~~~

### evals/export-edit-import-vba/graders/06-thisworkbook-updated-in-place.md

<!-- excel-com-file evals/export-edit-import-vba/graders/06-thisworkbook-updated-in-place.md 0ea89e34e2aece1e69051ba095f9a3238d87ab00060347d3eb53fcf841c75a5d -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

取り込み直した後の calc.xlsm の ThisWorkbook モジュールの BookVersion が "1.1" を返すコードになっており、"1.0" を返すコードは残っていない
~~~~~

### evals/export-edit-import-vba/graders/07-sheet-module-and-values.md

<!-- excel-com-file evals/export-edit-import-vba/graders/07-sheet-module-and-values.md 51a8871b0399c2b881e7ff4fe8b62928c14d2d89dc83957731f8b0b5f15771e8 -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

取り込み後に開き直した calc.xlsm の先頭シートで、元の文書モジュールに `Public SheetLabel As String` がある。A1 は 2、B1 の数式は `=A1*3`、計算結果は 6。
~~~~~

### evals/export-edit-import-vba/prompt.md

<!-- excel-com-file evals/export-edit-import-vba/prompt.md df08df113cdce2f7df2c2cb76dfc93dd2bbbef3a1773d4bcdfc61cd06ccfe252 -->
~~~~~markdown
---
tags: [skill-creator]
max_turns: 40
timeout_seconds: 1200
allowed_tools: [Read, Glob, Grep, Skill, Write, Edit, Bash]
---

カレントディレクトリに検証用の calc.xlsm を新しく作って、標準モジュール Util に「Public Function AddNums(a As Long, b As Long) As Long」（a+b を返す）を、ThisWorkbook に「Public Function BookVersion() As String」（"1.0" を返す）を入れて。そのあと VBA を書き出して、Util.bas の関数の上に「' 2 つの整数を足す」という日本語コメントを足し、ThisWorkbook の BookVersion が返す値を "1.1" に変えてから、ブックに取り込み直してほしい。このリポジトリの .bas のルールは守ってね。

書き出し先は `src/excel/calc/vba/` とし、`standard/`、`classes/`、`documents/` に分けて。書き出す前に、クラスモジュール CalcOptions に `Public Precision As Long`、先頭シートのモジュールに `Public SheetLabel As Long` を追加して。書き出した両方の .cls を編集して、型を String にしてから一緒に取り込んで。先頭シートの A1 に 2、B1 に =A1*3 を設定し、取り込み後に開き直して数式と計算結果を確認して。
~~~~~

### evals/jp-macro-ribbon-button/case.yaml

<!-- excel-com-file evals/jp-macro-ribbon-button/case.yaml bb7c887449ea918065c385fca5608c187d735333068bda7a71b26664c4a19b58 -->
~~~~~yaml
schema_version: "1.1"
name: jp-macro-ribbon-button
expected_outcome: |
  demo.xlsm が作られ、リボンの「ツール」タブに「完了通知」ボタンがある。
  ボタンの onAction は標準モジュールのマクロで、「処理が完了しました」を化けずに MsgBox で出す。
  Excel で開き直してリボンが読み込まれたことを確かめ、ボタンを実際に押してメッセージが出ることを観測する。
  作業後に Excel のプロセスが残っていない。
~~~~~

### evals/jp-macro-ribbon-button/graders/01-ribbon-part-and-relationship.md

<!-- excel-com-file evals/jp-macro-ribbon-button/graders/01-ribbon-part-and-relationship.md 34d963cccf6f3fde6fca03a2ef81b5a914d0d0abb22f514e15050e62a03af87b -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

demo.xlsm の _rels/.rels に customUI/customUI14.xml への ui/extensibility の関連付けがあり、customUI14.xml に label が「ツール」の tab と「完了通知」の button がある
~~~~~

### evals/jp-macro-ribbon-button/graders/02-callback-japanese-intact.md

<!-- excel-com-file evals/jp-macro-ribbon-button/graders/02-callback-japanese-intact.md d8d2a683583a46ac04a0b3dda69f95b1aabd44a2eccf372a8f30e754c4844805 -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

demo.xlsm の VBA 標準モジュールに、引数 control As IRibbonControl を取り「処理が完了しました」を MsgBox で表示する Sub があり、その文字列が化けていない（CodeModule から読み戻して一致する）
~~~~~

### evals/jp-macro-ribbon-button/graders/03-ribbon-load-verified.md

<!-- excel-com-file evals/jp-macro-ribbon-button/graders/03-ribbon-load-verified.md 25f21816cd97d83977027f96c32900af7dde81047e3b7fbe48d5d83f11e50893 -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

実行の中で demo.xlsm を Excel で開き直し、リボンが読み込まれたこと（onLoad の実行などの観測）を確かめている
~~~~~

### evals/jp-macro-ribbon-button/graders/04-click-verified-by-observation.md

<!-- excel-com-file evals/jp-macro-ribbon-button/graders/04-click-verified-by-observation.md 0cafe2e284d78f4a1d14a99e698328d77bf1466b543bed6e6871a1a16e375f0e -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

実行の中で「完了通知」ボタンを実際に押し、「処理が完了しました」を含むダイアログが出たことを観測している。押していないのに動作を確認済みと報告していない
~~~~~

### evals/jp-macro-ribbon-button/graders/05-no-excel-left.md

<!-- excel-com-file evals/jp-macro-ribbon-button/graders/05-no-excel-left.md 5ccff1bd657c0a71f5075b4e95d2e44f65cf4df002f5d68d26ef2ef9e46f3b5d -->
~~~~~markdown
---
type: llm
focus: trace
weight: 1
---

作業の終了時点で、実行中に起動した EXCEL.EXE が残っていない
~~~~~

### evals/jp-macro-ribbon-button/prompt.md

<!-- excel-com-file evals/jp-macro-ribbon-button/prompt.md 9e0dbd7b905734e3639c0005ea74e034922784638b4c7395d6f0224ca9781246 -->
~~~~~markdown
---
tags: [skill-creator]
max_turns: 40
timeout_seconds: 1200
allowed_tools: [Read, Glob, Grep, Skill, Write, Edit, Bash]
---

カレントディレクトリに demo.xlsm という新しいマクロ有効ブックを作って。リボンに「ツール」タブを足して、そこに「完了通知」ボタンを置いてほしい。ボタンを押したら「処理が完了しました」とメッセージボックスで出すマクロにして。作ったらちゃんと動くか確認しておいて。
~~~~~

### references/userform.md

<!-- excel-com-file references/userform.md 24d0fe91d5f224257aa3b295085bb6cafada2f68fcc43d126cfca992b056c41d -->
~~~~~markdown
# UserForm の作成と編集

## デザインと初期化の分担

フォームの固定レイアウトは、VBE でフォームを開いた時点で確認できるようにデザイナーへ保存する。タイトル、ラベル、固定のボタンやチェック欄、その配置・フォント・タブ順が該当する。旧コントロールを隠して固定レイアウトを作り直す処理がある場合は、参照とイベントを確認したうえでデザイナー側の配置に移す。[Microsoft Forms の Add メソッド](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms)

`UserForm_Initialize` には、シートの表示状態をチェック欄に反映する処理など、フォームを開いた時に決まる状態の初期化を置く。入力データによって項目数や種類が変わる場合は、必要なコントロールを実行時に追加できる。[Initialize イベント](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/initialize-event)

## 作業コピーから正本へ反映

1. [作業前の確認](../SKILL.md#作業前の確認)に従い、対象ブックを閉じた状態でバックアップと作業コピーを作る。既存フォームは作業コピーから別の場所へ書き出し、正本と照合する。VBE の再書き出しでは、先頭付近の空行や識別子の大文字・小文字、`.frx` のバイト列が変わることがある。差が出たらコードの意味を確認し、デザインは正本の `.frm` と `.frx` を別のコピーに取り込んで `Designer` のプロパティを比較する。残すべき変更を判断できない場合は、上書き前に確認する。
2. 作業コピーを `Use-ExcelWorkbook` で開き、既存の `VBComponent` を取得する。新規フォームは `VBComponents.Add(3)` で作る。フォームの `Caption`、`Width`、`Height` は `VBComponent.Properties.Item(...).Value`、コントロールは `VBComponent.Designer.Controls` から設定できる。フレーム内のコントロールはそのフレームの `Controls` に追加する。既存のコントロールを改名・削除する前に、イベントハンドラー名とコードからの参照を調べる。`Caption` をシート名などのキーに使っている場合は、その対応も確かめる。
3. 編集したコピーを保存し、`Export-Vba.ps1` で別の作業先に書き出す。対象フォームの `.frm` と `OleObjectBlob` が参照する同名の `.frx` を一組として扱う。書き出し前後でイベントコードを比較し、依頼に必要な初期化やイベントの差分だけを `.frm` に加える。イベントコードだけの変更では、既存のフォーム定義と `.frx` を使う。[VBA の編集ルール](vba.md)
4. 対象フォームの一組だけを `forms/` に置き、その親を `Import-Vba.ps1 -SourceDir` に渡して、編集前のブックの別コピーへ取り込む。次の検証が通った組を正本の `vba/forms/` と作業ブックへ反映する。

### PowerShell でフォントを設定できない場合

`Designer` から取得したコントロールに `.Font.Name` を代入したとき、PowerShell 5.1 と 7 で `NullReferenceException` が出た例がある。同じ代入を繰り返さず、作業コピーで VBE のプロパティを編集するか、`Designer` の `Font` プロパティを操作する一時 VBA を使う。一時 VBA を使う場合は[マクロが動くことへの注意](../SKILL.md#マクロが動くことへの注意)に従い、書き出し・保存前に一時モジュールを取り除く。

## 取り込み後の検証

- **デザイン時**は、保存したブックをマクロ無効のまま開き直し、`VBComponent.Designer` から変更したプロパティを読み戻す。フォームとコントロールの名前、表示文言、位置・大きさ、フォント、タブ順、フレーム内の構成を変更内容に応じて確認する。VBE の見た目が依頼の中心なら、可能な環境ではデザイナー上の表示も確認する。
- **実行時**は、許可された範囲でフォームをロードし、変更した初期化とイベントを確認する。たとえばチェック欄がシートの表示状態に連動する場合は、作業コピーで表示状態を変えてからフォームを読み直す。既存のボタンやマクロの実行は[マクロが動くことへの注意](../SKILL.md#マクロが動くことへの注意)に従う。デザインと実行時のどちらを確認したかを報告する。
- **ブック全体**は、フォームだけを変更した場合、`.xlsm` の ZIP 内で `xl/vbaProject.bin` 以外の差分を確認する。Excel の保存で、無関係な XML や作業先の絶対パスが書き込まれた実測がある。同じ未署名ブックのコピーで、フォームの変更が VBA プロジェクト内で完結する場合は、元のブックのコピーに Excel が保存した `xl/vbaProject.bin` だけを移せる。その場合も他の ZIP エントリが元と一致することを調べ、完成候補を Excel で開き直してフォームを検証する。
~~~~~

### references/vba.md

<!-- excel-com-file references/vba.md 64a562cffe4ffdb146a6d7c40505179e4f85b3ffdbdcf1fd211f0f6cc82a0e99 -->
~~~~~markdown
# VBA の編集ルール

## 保存形式とモジュール属性

- `.bas`、`.cls`、`.frm` の正本は [AGENTS.md](../../../../AGENTS.md) に従い、UTF-8（BOM なし）・LF で保存する。VBE との文字コード・改行の変換は入出力スクリプトに任せる。
- 書き出したファイルを編集する時は、`Attribute VB_Name` などのモジュール属性を保持する。新しいクラスの `.cls` には、VBE が取り込めるクラスヘッダーと属性を含める。
- 文書モジュールの属性は書き出したファイルに保持する。シートとブックへの反映では、既存モジュールのコード部分を差し替える。

## 識別子

- モジュール名（`Attribute VB_Name`）は31文字以内にする。プロシージャや変数などの一般識別子は255文字以内にする。
- 識別子は文字で始める。空白、ピリオド、`!`、`@`、`&`、`$`、`#` を含めない。予約語と同一スコープでの名前の衝突を避ける。
- `Option Explicit` はモジュールの宣言部に置く。

## 変更前の確認

- 公開プロシージャ、マクロの入口、シート名やモジュール名を変更する前に、他のモジュール、シートの数式、リボン XML からの参照を確認する。
- `src/excel/<ブック名>/` の正本とブック内のコードの同期が不明な場合は、ブックから別の作業先に書き出して差分を確認する。どの変更を残すか判断できない場合は、反映前に確認する。
~~~~~

### scripts/Add-CustomUI.ps1

<!-- excel-com-file scripts/Add-CustomUI.ps1 2e113d488939a2b619c60664aedcc8ce95bcc6539717c80a17d7747609d879f0 -->
~~~~~powershell
<#
.SYNOPSIS
Embeds or replaces customUI/customUI14.xml (Office 2010+ RibbonX) in a closed .xlsm package.

.DESCRIPTION
Excel's COM API cannot define a ribbon, so the XML is written straight into the OOXML zip.
The workbook must be closed. If customUI14.xml already exists its content is replaced;
the package-level relationship in _rels/.rels is added only when missing.
The updated package is checked in a sibling temporary file before replacing the workbook.
The new part relies on the default "xml" content type in [Content_Types].xml; the script stops if it is missing.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$XmlPath
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$partName = 'customUI/customUI14.xml'
$relType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
$relNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
$utf8 = New-Object System.Text.UTF8Encoding($false)

$Path = (Resolve-Path -LiteralPath $Path).ProviderPath
# Parse first so that malformed XML never reaches the workbook.
[xml]$ui = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $XmlPath).ProviderPath, $utf8)

function Read-PackageXml($Zip, [string]$Part) {
    $entry = $Zip.GetEntry($Part)
    if (!$entry) { throw "Missing package part: $Part" }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Write-PackageXml($Zip, [string]$Part, [xml]$Document) {
    $old = $Zip.GetEntry($Part)
    if ($old) { $old.Delete() }
    $writer = New-Object System.IO.StreamWriter($Zip.CreateEntry($Part).Open(), $utf8)
    try { $Document.Save($writer) } finally { $writer.Dispose() }
}

function Get-RibbonRelationship([xml]$Relationships) {
    $Relationships.DocumentElement.ChildNodes |
        Where-Object { $_.Type -eq $relType -and $_.Target.TrimStart('/') -eq $partName }
}

$tempPath = Join-Path ([IO.Path]::GetDirectoryName($Path)) ('.customui-' + [guid]::NewGuid() + '.tmp')
try {
    [IO.File]::Copy($Path, $tempPath)
    $zip = [System.IO.Compression.ZipFile]::Open($tempPath, 'Update')
    try {
        $ct = Read-PackageXml $zip '[Content_Types].xml'
        if (-not ($ct.Types.Default | Where-Object { $_.Extension -eq 'xml' })) {
            throw '[Content_Types].xml has no Default for the "xml" extension'
        }

        $replacing = $null -ne $zip.GetEntry($partName)
        $rels = Read-PackageXml $zip '_rels/.rels'
        $exists = @(Get-RibbonRelationship $rels).Count -gt 0
        Write-PackageXml $zip $partName $ui
        if (!$exists) {
            $node = $rels.CreateElement('Relationship', $relNs)
            $id = 'rIdCustomUI14'
            if ($id -in $rels.DocumentElement.ChildNodes.Id) { $id += [guid]::NewGuid().ToString('N') }
            $node.SetAttribute('Id', $id)
            $node.SetAttribute('Type', $relType)
            $node.SetAttribute('Target', $partName)
            [void]$rels.DocumentElement.AppendChild($node)
            Write-PackageXml $zip '_rels/.rels' $rels
        }
    }
    finally { $zip.Dispose() }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
    try {
        $written = Read-PackageXml $zip $partName
        $rels = Read-PackageXml $zip '_rels/.rels'
        if ($written.DocumentElement.OuterXml -cne $ui.DocumentElement.OuterXml -or
            @(Get-RibbonRelationship $rels).Count -ne 1) {
            throw 'Could not verify the custom UI part and its package relationship'
        }
    }
    finally { $zip.Dispose() }

    [IO.File]::Replace($tempPath, $Path, [NullString]::Value)
    "customUI14.xml: " + $(if ($replacing) { 'replaced' } else { 'added' })
    "relationship: " + $(if ($exists) { 'kept' } else { 'added' })
}
finally {
    if ([IO.File]::Exists($tempPath)) { [IO.File]::Delete($tempPath) }
}
~~~~~

### scripts/ExcelCom.psm1

<!-- excel-com-file scripts/ExcelCom.psm1 10b84cab0da51ba507c3eae4f7b94b9f3db81f8cf4f45d479a5e546646f95b10 -->
~~~~~powershell
<#
.SYNOPSIS
Runs Excel through COM in a dedicated instance and always cleans up.

.DESCRIPTION
Use-ExcelWorkbook opens an existing workbook and invokes a script block with the Excel.Application
and Workbook objects. New-ExcelWorkbook creates a new empty workbook file.

Use-ExcelWorkbook disables VBA macros and workbook events by default and opens without updating
external workbook links. Use -EnableMacros for an intentional macro test. Event handlers such as
Workbook_Open also require -EnableEvents.

On exit the workbook is closed without saving (unless -Save was given and the block completed),
Excel is asked to quit, and the COM objects are released. Each step is attempted even if an earlier
one fails. If this instance's process is still alive afterwards, its PID is reported as a warning;
it is not stopped automatically.
#>

# Add-Type fails if the type already exists, e.g. when the module is imported again with -Force.
if (-not ('ExcelCom.Native' -as [type])) {
    Add-Type -Namespace ExcelCom -Name Native -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@
}

function Get-ExcelProcessId($Xl) {
    # Hwnd is available even for a hidden instance.
    $processId = [uint32]0
    [void][ExcelCom.Native]::GetWindowThreadProcessId([IntPtr]$Xl.Hwnd, [ref]$processId)
    [int]$processId
}

function Start-ExcelSession([bool]$Visible, [bool]$EnableEvents) {
    $xl = New-Object -ComObject Excel.Application
    $session = @{ Xl = $xl; ProcessId = (Get-ExcelProcessId $xl) }
    $xl.Visible = $Visible
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $EnableEvents
    $session
}

function Stop-ExcelSession($Session, $Workbook) {
    if ($null -ne $Workbook) {
        try { $Workbook.Close($false) } catch { Write-Warning ("Workbook.Close failed: " + $_.Exception.Message) }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) } catch { }
    }
    try { $Session.Xl.Quit() } catch { Write-Warning ("Excel.Quit failed: " + $_.Exception.Message) }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Session.Xl) } catch { }
    $Session.Xl = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Quit returns before the process exits; give it a few seconds before reporting.
    for ($i = 0; $i -lt 10; $i++) {
        if (-not (Get-Process -Id $Session.ProcessId -ErrorAction SilentlyContinue)) { return }
        Start-Sleep -Milliseconds 500
    }
    Write-Warning ("The Excel instance started by this call (PID " + $Session.ProcessId +
        ") is still running. Stop only this PID if needed.")
}

function Use-ExcelWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        # Show the window. Needed when the ribbon (customUI) must be built, e.g. to click a button.
        [switch]$Visible,
        # Save the workbook after the script block completes without throwing.
        [switch]$Save,
        # Let Workbook_Open and other event handlers run.
        [switch]$EnableEvents,
        [switch]$EnableMacros
    )
    $Path = (Resolve-Path -LiteralPath $Path).ProviderPath
    $session = Start-ExcelSession ([bool]$Visible) ([bool]$EnableEvents)
    $wb = $null
    try {
        if (!$EnableMacros) {
            $session.Xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        }
        $wb = $session.Xl.Workbooks.Open($Path, 0)  # UpdateLinks = 0
        & $ScriptBlock $session.Xl $wb
        if ($Save) { $wb.Save() }
    }
    finally {
        Stop-ExcelSession $session $wb
    }
}

function New-ExcelWorkbook {
    [CmdletBinding()]
    param(
        # Target file. The format follows the extension: .xlsm (macro-enabled) or .xlsx.
        [Parameter(Mandatory)][string]$Path
    )
    $Path = [IO.Path]::GetFullPath($Path)
    if (Test-Path -LiteralPath $Path) { throw "Already exists: $Path" }
    $format = switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.xlsm' { 52 }   # xlOpenXMLWorkbookMacroEnabled
        '.xlsx' { 51 }   # xlOpenXMLWorkbook
        default { throw "Unsupported extension: $Path" }
    }
    $session = Start-ExcelSession $false $false
    $wb = $null
    try {
        $wb = $session.Xl.Workbooks.Add()
        $wb.SaveAs($Path, $format)
    }
    finally {
        Stop-ExcelSession $session $wb
    }
    $Path
}

Export-ModuleMember -Function Use-ExcelWorkbook, New-ExcelWorkbook, Get-ExcelProcessId
~~~~~

### scripts/Export-Vba.ps1

<!-- excel-com-file scripts/Export-Vba.ps1 3d5cbc548992edd2839a6018beebc45ea6325fa159a36cbc0f87c9e827f682aa -->
~~~~~powershell
<#
.SYNOPSIS
Exports workbook VBA modules and UserForms using the repository's layout.

.DESCRIPTION
Standard modules -> <Name>.bas, class modules and document modules (ThisWorkbook, sheets) -> <Name>.cls.
The respective files go into standard/, classes/, and documents/ under OutDir.
UserForms go into forms/ as .frm with the accompanying .frx if exported by VBE.
Document modules without code are skipped.

Text exports are converted from CP932/CRLF to UTF-8 without BOM and LF. Binary .frx files are copied unchanged.

Files in OutDir that were not written by this run are listed as a warning (a module removed from the
workbook, or a document module whose code is now empty), because importing them again would bring
the old code back.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$OutDir
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelCom.psm1') -Force

$cp932 = [Text.Encoding]::GetEncoding(932, [Text.EncoderFallback]::ExceptionFallback,
    [Text.DecoderFallback]::ExceptionFallback)
$utf8 = New-Object System.Text.UTF8Encoding($false)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$OutDir = (Resolve-Path -LiteralPath $OutDir).ProviderPath
$folders = @{ 1 = 'standard'; 2 = 'classes'; 3 = 'forms'; 100 = 'documents' }
if (Get-ChildItem -LiteralPath $OutDir -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm', '.frx' }) {
    throw "Place VBA files in standard/, classes/, documents/, or forms/, not directly in $OutDir"
}
$outputDirectories = @(foreach ($folder in $folders.Values) {
    $directory = Join-Path $OutDir $folder
    New-Item -ItemType Directory -Force -Path $directory | Out-Null
    $directory
})
$tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('vba-export-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $tmpDir | Out-Null

try {
    $written = Use-ExcelWorkbook -Path $Path -ScriptBlock {
        param($xl, $wb)
        foreach ($c in $wb.VBProject.VBComponents) {
            if ($c.Type -eq 1) { $ext = 'bas' }
            elseif ($c.Type -eq 2) { $ext = 'cls' }
            elseif ($c.Type -eq 3) { $ext = 'frm' }
            elseif ($c.Type -eq 100) {
                if ($c.CodeModule.CountOfLines -eq 0) { continue }
                $ext = 'cls'
            }
            else {
                Write-Warning ("Skipped " + $c.Name + " (component type " + $c.Type + " is not supported)")
                continue
            }
            $file = $c.Name + '.' + $ext
            $tmp = Join-Path $tmpDir $file
            $c.Export($tmp)
            try { $text = $cp932.GetString([IO.File]::ReadAllBytes($tmp)) }
            catch { throw ("Cannot decode Excel export as CP932: " + $file) }
            $text = $text -replace "`r`n?", "`n"
            $directory = Join-Path $OutDir $folders[[int]$c.Type]
            $outputFile = Join-Path $directory $file
            if ($ext -eq 'frm') {
                $sidecarName = $c.Name + '.frx'
                $sidecar = Join-Path $tmpDir $sidecarName
                if (Test-Path -LiteralPath $sidecar -PathType Leaf) {
                    $outputSidecar = Join-Path $directory $sidecarName
                    [IO.File]::Copy($sidecar, $outputSidecar, $true)
                    $outputSidecar
                }
            }
            [IO.File]::WriteAllText($outputFile, $text, $utf8)
            $outputFile
        }
    }
}
finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

$written | ForEach-Object { "exported: " + $_.Substring($OutDir.Length).TrimStart('\', '/') }
$stale = @(Get-ChildItem -LiteralPath $outputDirectories -File |
    Where-Object { $_.Extension -in '.bas', '.cls', '.frm', '.frx' -and $written -notcontains $_.FullName })
if ($stale.Count -gt 0) {
    Write-Warning ("Present in OutDir but not written by this export (module removed, or document module now empty): " +
        (($stale | ForEach-Object { $_.FullName.Substring($OutDir.Length).TrimStart('\', '/') }) -join ', '))
}
~~~~~

### scripts/Import-Vba.ps1

<!-- excel-com-file scripts/Import-Vba.ps1 2bef8bf10520bce0cfb5f658d0ac53157affa8b472e1aaca65cd996d46f907d8 -->
~~~~~powershell
<#
.SYNOPSIS
Imports UTF-8 (no BOM), LF .bas / .cls / .frm files and saves the workbook.

.DESCRIPTION
- All sources are checked for lossless CP932 conversion before opening the workbook.
- Temporary files use CP932 and CRLF. VBComponents.Import misreads an LF-only .cls as a standard module.
- An existing standard module, class module, or UserForm with the same name and type is replaced.
- A UserForm's referenced .frx is copied unchanged next to the temporary .frm before import.
- A document module (ThisWorkbook, sheets) cannot be imported. Its code is replaced instead:
  the class header and the leading Attribute lines are dropped and the rest is written to the
  existing module. Attribute lines inside procedures are not handled.
- The standard/, classes/, documents/, and forms/ directories determine the component type.
- Every existing component must match the source type before any code is changed.
- After import, each module's name and type are checked against the file.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$SourceDir
)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'ExcelCom.psm1') -Force

$cp932 = [Text.Encoding]::GetEncoding(932, [Text.EncoderFallback]::ExceptionFallback,
    [Text.DecoderFallback]::ExceptionFallback)
if (Get-ChildItem -LiteralPath $SourceDir -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm', '.frx' }) {
    throw "Place VBA files in standard/, classes/, documents/, or forms/, not directly in $SourceDir"
}
$folderTypes = @{ standard = 1; classes = 2; forms = 3; documents = 100 }
$extensions = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm'; 100 = '.cls' }
$sourceFolders = @(Get-ChildItem -LiteralPath $SourceDir -Directory |
    Where-Object { $folderTypes.ContainsKey($_.Name) })
$files = @(foreach ($folder in $sourceFolders) {
    Get-ChildItem -LiteralPath $folder.FullName -File | Where-Object { $_.Extension -in '.bas', '.cls', '.frm' }
})
if ($files.Count -eq 0) { throw "No .bas / .cls / .frm files in VBA component directories under $SourceDir" }

$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
$sources = foreach ($f in $files) {
    $bytes = [IO.File]::ReadAllBytes($f.FullName)
    try { $text = $strictUtf8.GetString($bytes) }
    catch { throw ("Invalid UTF-8 in " + $f.FullName) }
    if ($text.StartsWith([string][char]0xFEFF, [StringComparison]::Ordinal) -or $text.Contains("`r")) {
        throw ("Expected UTF-8 without BOM and LF line endings in " + $f.FullName)
    }
    $text = $text.Replace("`n", "`r`n")
    try { $importBytes = $cp932.GetBytes($text) }
    catch { throw ("Cannot encode as CP932: " + $f.FullName) }
    if (![string]::Equals($cp932.GetString($importBytes), $text, [StringComparison]::Ordinal)) {
        throw ("Cannot round-trip through CP932: " + $f.FullName)
    }
    $m = [regex]::Match($text, '(?m)^Attribute VB_Name = "([^"]+)"\r$')
    if (-not $m.Success) { throw ("No Attribute VB_Name in " + $f.Name) }
    $name = $m.Groups[1].Value
    $type = $folderTypes[$f.Directory.Name]
    if ($f.Extension -ne $extensions[$type] -or $f.BaseName -ne $name) {
        throw ("Unexpected extension or VB_Name in " + $f.FullName)
    }
    $sidecar = $null
    if ($type -eq 3) {
        $designer = $text.Substring(0, $m.Index)
        if ($designer -match '(?m)^[ \t]*OleObjectBlob[ \t]*=') {
            $blob = [regex]::Match($designer, '(?m)^[ \t]*OleObjectBlob[ \t]*=[ \t]*"([^"]+)":[0-9A-Fa-f]+[ \t]*\r$')
            $sidecarName = $f.BaseName + '.frx'
            if (!$blob.Success -or $blob.Groups[1].Value -ne $sidecarName) {
                throw ("Expected a reference to " + $sidecarName + " in " + $f.FullName)
            }
            $sidecar = Join-Path $f.DirectoryName $sidecarName
            if (!(Test-Path -LiteralPath $sidecar -PathType Leaf)) {
                throw ("Missing UserForm resource: " + $sidecar)
            }
        }
    }
    [pscustomobject]@{
        File = $f.Name
        Name = $name
        Text = $text
        ImportBytes = $importBytes
        Sidecar = $sidecar
        Type = $type
    }
}

$duplicate = $sources | Group-Object Name | Where-Object Count -gt 1
if ($duplicate) { throw ("Duplicate VB_Name: " + ($duplicate.Name -join ', ')) }

$tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('vba-import-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $tmpDir | Out-Null
try {
    Use-ExcelWorkbook -Path $Path -Save -ScriptBlock {
        param($xl, $wb)
        $components = $wb.VBProject.VBComponents
        $existingByName = @{}
        foreach ($component in $components) { $existingByName[$component.Name] = $component }
        foreach ($s in $sources) {
            $existing = $existingByName[$s.Name]
            if ($null -ne $existing -and $existing.Type -ne $s.Type) {
                throw ("Component type mismatch for " + $s.Name + ": workbook " + $existing.Type + ", source " + $s.Type)
            }
            if ($s.Type -eq 100 -and $null -eq $existing) {
                throw ("No document module named " + $s.Name + " in the workbook")
            }
        }
        foreach ($s in $sources) {
            $existing = $existingByName[$s.Name]
            if ($s.Type -eq 100) {
                $lines = $s.Text -split "`r`n"
                $i = 0
                if ($lines[0] -eq 'VERSION 1.0 CLASS') {
                    while ($i -lt $lines.Count -and $lines[$i] -ne 'END') { $i++ }
                    $i++
                }
                while ($i -lt $lines.Count -and $lines[$i] -like 'Attribute *') { $i++ }
                $code = ''
                if ($i -lt $lines.Count) { $code = ($lines[$i..($lines.Count - 1)] -join "`r`n").TrimEnd("`r", "`n") }
                $cm = $existing.CodeModule
                if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
                if ($code.Length -gt 0) { $cm.AddFromString($code) }
                "replaced document module: " + $s.Name
                continue
            }
            $tmp = Join-Path $tmpDir $s.File
            [IO.File]::WriteAllBytes($tmp, $s.ImportBytes)
            if ($s.Sidecar) {
                [IO.File]::Copy($s.Sidecar, [IO.Path]::ChangeExtension($tmp, '.frx'))
            }
            if ($null -ne $existing) { $components.Remove($existing) }
            $c = $components.Import($tmp)
            if ($c.Name -ne $s.Name -or $c.Type -ne $s.Type) {
                throw ("Imported " + $s.File + " as " + $c.Name + " (type " + $c.Type + "), expected " +
                    $s.Name + " (type " + $s.Type + ")")
            }
            "imported: " + $s.Name
        }
    }
}
finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}
~~~~~

### scripts/Invoke-RibbonButton.ps1

<!-- excel-com-file scripts/Invoke-RibbonButton.ps1 e63803a2e3987ff11a5d73f28ebd27abb3c0ee1c6d6a67b271ebf2eb2ee27671 -->
~~~~~powershell
<#
.SYNOPSIS
Clicks a custom ribbon button through UI Automation and reports the dialogs it shows.

.DESCRIPTION
Opens the workbook in a visible Excel, selects the tab, invokes the button, and handles the dialogs
(#32770, e.g. MsgBox) that appear afterwards: the names of each dialog's elements are recorded and the
dialog is closed with WindowPattern.Close (the same as the x button). Only this check has been done with
a visible window; whether the ribbon is built for a hidden one has not been checked. The workbook is
closed without saving. Workbook_Open does not run (see ExcelCom.psm1), the button's own macro does.
The tab and the button must be uniquely named in their respective search scopes.

In a MsgBox both the message and the buttons are exposed as Pane elements, so they cannot be told
apart by control type. Each entry of Dialogs lists them all, e.g. ["OK", "<message>"].

If a dialog cannot be closed (a Yes/No MsgBox has no Cancel and ignores the x button), or the click
call does not return in time, the Excel instance started by this script is stopped by its PID and the
result records the error. Other Excel processes are never touched.

The click call returns without waiting for the macro to finish. A macro that keeps running without a
dialog (e.g. Application.Wait) is waited for when the workbook is closed; a macro that never ends
makes this script hang there.

The result is written as UTF-8 JSON to -OutFile (stdout of powershell.exe garbles Japanese in Git Bash).
Exit code is non-zero when the click failed.
#>
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$TabLabel,
    [Parameter(Mandatory)][string]$ButtonLabel,
    [Parameter(Mandatory)][string]$OutFile,
    # How long to wait for the tab / button to appear.
    [int]$FindTimeoutSeconds = 15,
    # How long to wait for each dialog. No dialog within this time is not an error.
    [int]$DialogTimeoutSeconds = 5,
    # How long to wait for the click call to return after the last dialog.
    [int]$ActionTimeoutSeconds = 10,
    # Safety limit on the number of dialogs handled for one click.
    [int]$MaxDialogs = 5
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Import-Module (Join-Path $PSScriptRoot 'ExcelCom.psm1') -Force

$AE = [System.Windows.Automation.AutomationElement]
$Scope = [System.Windows.Automation.TreeScope]
$ControlType = [System.Windows.Automation.ControlType]

function Find-Element($Root, $Condition, [double]$TimeoutSeconds) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $el = $Root.FindFirst($Scope::Descendants, $Condition)
        if ($el) { return $el }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)
    $null
}

function Get-DialogNames($Dialog) {
    $names = @()
    foreach ($e in $Dialog.FindAll($Scope::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
        if ($e.Current.Name) { $names += $e.Current.Name }
    }
    , $names
}

function Find-UniqueElement($Root, $Condition, [double]$TimeoutSeconds, [string]$Description) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $matches = $Root.FindAll($Scope::Descendants, $Condition)
        if ($matches.Count -gt 1) { throw "Ambiguous ${Description}: $($matches.Count) matching elements" }
        if ($matches.Count -eq 1) { return $matches.Item(0) }
        Start-Sleep -Milliseconds 300
    } while ((Get-Date) -lt $deadline)
    throw "Not found: $Description"
}

function New-NamedControlCondition([string]$Name, $Type) {
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)
    $typeCondition = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $Type)
    [System.Windows.Automation.AndCondition]::new([System.Windows.Automation.Condition[]]@($nameCondition, $typeCondition))
}

$result = [ordered]@{ TabFound = $false; ButtonFound = $false; Clicked = $false; Dialogs = @(); Error = $null }
try {
    Use-ExcelWorkbook -Path $Path -Visible -EnableMacros -ScriptBlock {
        param($xl, $wb)
        $excelPid = Get-ExcelProcessId $xl
        $win = $AE::FromHandle([IntPtr]$xl.Hwnd)

        $tab = Find-UniqueElement $win (New-NamedControlCondition $TabLabel $ControlType::TabItem) $FindTimeoutSeconds "tab '$TabLabel'"
        $result.TabFound = $true
        $tabRoot = [System.Windows.Automation.TreeWalker]::RawViewWalker.GetParent($tab)
        if (!$tabRoot -or $tabRoot.Current.ControlType -ne $ControlType::Tab) {
            throw "Ribbon tab container not found: $TabLabel"
        }
        $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()

        # Excel exposes the selected tab's controls in a sibling group, not under the tab item.
        $tabContent = Find-UniqueElement $tabRoot (New-NamedControlCondition $TabLabel $ControlType::Group) $FindTimeoutSeconds "tab content '$TabLabel'"
        $btn = Find-UniqueElement $tabContent (New-NamedControlCondition $ButtonLabel $ControlType::Button) $FindTimeoutSeconds "button '$ButtonLabel' in '$TabLabel'"
        $result.ButtonFound = $true

        # Invoke does not return while the macro shows a modal dialog, so click from another runspace.
        $ps = [PowerShell]::Create()
        [void]$ps.AddScript({ param($b)
            $ErrorActionPreference = 'Stop'
            $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }).AddArgument($btn)
        $async = $ps.BeginInvoke()

        $stuck = $null
        $dlgCond = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, '#32770')
        while ($result.Dialogs.Count -lt $MaxDialogs) {
            # A MsgBox shows up as a descendant of the Excel window in the UIA tree.
            $dlg = Find-Element $win $dlgCond $DialogTimeoutSeconds
            if (-not $dlg) { break }
            $names = Get-DialogNames $dlg
            $result.Dialogs += , ([ordered]@{ Title = $dlg.Current.Name; Names = $names })
            $handle = $dlg.Current.NativeWindowHandle
            try { $dlg.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() } catch { }
            # The next MsgBox may reuse the same window, so "closed" means the same window with the
            # same contents is no longer there.
            $closed = $false
            for ($i = 0; $i -lt 10 -and -not $closed; $i++) {
                Start-Sleep -Milliseconds 300
                $now = $win.FindFirst($Scope::Descendants, $dlgCond)
                $closed = (-not $now) -or ($now.Current.NativeWindowHandle -ne $handle) -or
                    ((Get-DialogNames $now) -join "`n") -ne ($names -join "`n")
            }
            if (-not $closed) { $stuck = 'A dialog could not be closed'; break }
        }
        if (-not $stuck -and -not $async.AsyncWaitHandle.WaitOne($ActionTimeoutSeconds * 1000)) {
            $stuck = 'The click call did not return in time'
        }
        if ($stuck) {
            # The instance is blocked in a modal state; COM calls to it would hang. Stop only this instance.
            # The runspace is left undisposed on purpose: Dispose waits for the blocked Invoke.
            Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            throw "$stuck. Stopped the Excel instance of this script (PID $excelPid)."
        }
        try {
            [void]$ps.EndInvoke($async)
            if ($ps.HadErrors) { throw ($ps.Streams.Error | Select-Object -First 1).Exception }
        }
        finally { $ps.Dispose() }
        $result.Clicked = $true
    }
}
catch {
    $result.Error = $_.Exception.Message
}
finally {
    $json = [pscustomobject]$result | ConvertTo-Json -Depth 5
    [IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}
if ($result.Error) { Write-Error $result.Error; exit 1 }
~~~~~

### scripts/Lock-VbaProject.ps1

<!-- excel-com-file scripts/Lock-VbaProject.ps1 0bfc779872b956793519cd10fd6f7ce6e60f1e504975a638f743d6786b165c1c -->
~~~~~powershell
# Locks a closed, disposable publish copy and verifies the saved project.
param([Parameter(Mandatory)][string]$Path)
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelCom.psm1')
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
if (-not ('VbaProjectLock.Native' -as [type])) {
    Add-Type -Namespace VbaProjectLock -Name Native -MemberDefinition @'
[DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
private static extern IntPtr SendMessageTimeout(
    IntPtr window, uint message, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);

[DllImport("user32.dll", SetLastError = true)]
public static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

public static long Send(IntPtr window, uint message, int wParam, IntPtr lParam)
{
    IntPtr result;
    // SMTO_ABORTIFHUNG | SMTO_ERRORONEXIT
    if (SendMessageTimeout(window, message, new IntPtr(wParam), lParam, 0x22, 2000, out result) == IntPtr.Zero)
        throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
    return result.ToInt64();
}
'@
}

$setProtection = {
    param($vbeHandle, $projectName, $excelPid, $commandReturned)
    $ErrorActionPreference = 'Stop'
    $element = [System.Windows.Automation.AutomationElement]
    $scope = [System.Windows.Automation.TreeScope]
    $dialogCondition = [System.Windows.Automation.PropertyCondition]::new($element::ClassNameProperty, '#32770')
    $dialogHandle = [IntPtr]::Zero
    $accepted = $false

    function Get-ControlHandle([string]$Id, [string]$ClassName) {
        $idCondition = [System.Windows.Automation.PropertyCondition]::new($element::AutomationIdProperty, $Id)
        $controls = @($dialog.FindAll($scope::Descendants, $idCondition) |
            Where-Object { $_.Current.ClassName -eq $ClassName })
        if ($controls.Count -ne 1 -or !$controls[0].Current.NativeWindowHandle) {
            throw "Expected one VBE protection control with ID $Id ($ClassName)."
        }
        [IntPtr]$controls[0].Current.NativeWindowHandle
    }

    try {
        $vbeWindow = $element::FromHandle($vbeHandle)
        $deadline = [DateTime]::UtcNow.AddSeconds(10)
        do {
            $matches = @($vbeWindow.FindAll($scope::Descendants, $dialogCondition) |
                Where-Object { $_.Current.ProcessId -eq $excelPid -and
                    $_.Current.Name.StartsWith($projectName + ' - ', [StringComparison]::Ordinal) })
            if ($matches.Count -gt 1) { throw 'Multiple VBE project properties dialogs found.' }
            if ($matches.Count -eq 1) { break }
        } while ([DateTime]::UtcNow -lt $deadline -and !$commandReturned.Wait(150))
        if ($matches.Count -ne 1) { throw 'VBE project properties dialog not found.' }
        $dialog = $matches[0]
        $dialogHandle = [IntPtr]$dialog.Current.NativeWindowHandle

        # The VBE exposes these legacy controls as panes. Use their Win32 IDs and messages.
        $tabs = Get-ControlHandle '12320' 'SysTabControl32'
        if ([VbaProjectLock.Native]::Send($tabs, 0x1304, 0, 0) -ne 2) {  # TCM_GETITEMCOUNT
            throw 'Unexpected VBE project properties tab layout.'
        }
        [void][VbaProjectLock.Native]::Send($tabs, 0x1330, 1, 0)  # TCM_SETCURFOCUS notifies the dialog
        $lock = Get-ControlHandle '5463' 'Button'
        $password = Get-ControlHandle '5461' 'Edit'
        $confirmation = Get-ControlHandle '5462' 'Edit'
        $ok = Get-ControlHandle '1' 'Button'

        # Use the Edit controls' input path so VBE receives the change notifications.
        $random = [Security.Cryptography.RandomNumberGenerator]::Create()
        $bytes = [byte[]]::new(8)
        try { $random.GetBytes($bytes) }
        finally { $random.Dispose() }
        $text = [BitConverter]::ToString($bytes).Replace('-', '')
        foreach ($field in $password, $confirmation) {
            [void][VbaProjectLock.Native]::Send($field, 0x00B1, 0, -1)  # EM_SETSEL
            foreach ($character in $text.ToCharArray()) {
                [void][VbaProjectLock.Native]::Send($field, 0x0102, [int]$character, 1)  # WM_CHAR
            }
            if ([VbaProjectLock.Native]::Send($field, 0x000E, 0, 0) -ne $text.Length) {  # WM_GETTEXTLENGTH
                throw 'VBE did not accept the generated password.'
            }
        }
        if ([VbaProjectLock.Native]::Send($lock, 0x00F0, 0, 0) -ne 1) {  # BM_GETCHECK
            [void][VbaProjectLock.Native]::Send($lock, 0x00F5, 0, 0)  # BM_CLICK
        }
        if ([VbaProjectLock.Native]::Send($lock, 0x00F0, 0, 0) -ne 1) {
            throw 'VBE did not select Lock project for viewing.'
        }
        [void][VbaProjectLock.Native]::Send($ok, 0x00F5, 0, 0)  # BM_CLICK
        if (!$commandReturned.Wait(5000)) {
            throw 'VBE did not accept the project protection settings.'
        }
        $accepted = $true
    }
    finally {
        if (!$accepted -and !$commandReturned.IsSet) {
            if ($dialogHandle -ne [IntPtr]::Zero) {
                [void][VbaProjectLock.Native]::PostMessage($dialogHandle, 0x0010, 0, 0)  # WM_CLOSE
            }
            if (!$commandReturned.Wait(3000)) {
                # Only the dedicated Excel process can be blocked by our modal command.
                Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

Use-ExcelWorkbook -Path $Path -Visible -Save -ScriptBlock {
    param($xl, $wb)
    $project = $vbe = $vbeWindow = $commandBars = $properties = $null
    try {
        $project = $wb.VBProject
        if ($project.Protection -ne 0) { throw 'The publish copy must have an unlocked VBA project.' }
        $vbe = $xl.VBE
        $vbeWindow = $vbe.MainWindow
        $vbeWindow.Visible = $true
        $vbe.ActiveVBProject = $project
        $commandBars = $vbe.CommandBars
        $properties = $commandBars.FindControl(1, 2578)  # msoControlButton, Project Properties
        if (!$properties -or !$properties.Enabled) { throw 'VBE Project Properties command is unavailable.' }

        # Execute blocks until the modal dialog closes, so operate the controls from a second runspace.
        $commandReturned = [Threading.ManualResetEventSlim]::new()
        $worker = [PowerShell]::Create()
        try {
            [void]$worker.AddScript($setProtection.ToString()).AddArgument([IntPtr]$vbeWindow.HWnd).
                AddArgument($project.Name).AddArgument((Get-ExcelProcessId $xl)).AddArgument($commandReturned)
            $operation = $worker.BeginInvoke()
            try { $properties.Execute() }
            finally {
                $commandReturned.Set()
                [void]$worker.EndInvoke($operation)
            }
        }
        finally {
            $worker.Dispose()
            $commandReturned.Dispose()
        }
    }
    finally {
        # Release the VBE objects before Use-ExcelWorkbook closes the workbook and quits Excel.
        foreach ($reference in $properties, $commandBars, $vbeWindow, $vbe, $project) {
            if ($null -ne $reference) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($reference)
            }
        }
    }
}

Use-ExcelWorkbook -Path $Path -ScriptBlock {
    param($xl, $wb)
    if ($wb.VBProject.Protection -ne 1) {  # vbext_pp_locked
        throw 'The saved VBA project is not locked. The publish copy was rejected.'
    }
}
~~~~~

### scripts/Publish-Workbook.ps1

<!-- excel-com-file scripts/Publish-Workbook.ps1 9f90e2895a3b39964350a4f786948de86912301b61489a512f2f6bd3944cad71 -->
~~~~~powershell
<#
.SYNOPSIS
Updates workbook.xlsm from its VBA and ribbon sources and publishes a VBA-locked copy.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [ValidateNotNullOrEmpty()][string]$FileName,
    [ValidateNotNullOrEmpty()][string]$OutDir,
    [switch]$Force
)
$ErrorActionPreference = 'Stop'

$workbook = Get-Item -LiteralPath $Path
if ($workbook -isnot [IO.FileInfo] -or $workbook.Name -ne 'workbook.xlsm') {
    throw "Expected a workbook.xlsm file: $Path"
}

$sourceDir = Join-Path $workbook.DirectoryName 'vba'
$ribbonPath = Join-Path $workbook.DirectoryName 'ribbon/customUI14.xml'
if (!(Test-Path -LiteralPath $sourceDir -PathType Container)) { throw "Missing VBA directory: $sourceDir" }
if (!(Test-Path -LiteralPath $ribbonPath -PathType Leaf)) { throw "Missing ribbon XML: $ribbonPath" }

if (!$PSBoundParameters.ContainsKey('FileName')) {
    $FileName = $workbook.Directory.Name + '.xlsm'
}
if ([IO.Path]::GetExtension($FileName) -ne '.xlsm' -or
    $FileName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) {
    throw "Expected a file name ending in .xlsm: $FileName"
}
if (!$PSBoundParameters.ContainsKey('OutDir')) {
    $OutDir = Join-Path $PSScriptRoot '../../../..'
}
$provider = $null
$drive = $null
$outputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
    $OutDir, [ref]$provider, [ref]$drive)
if ($provider.Name -ne 'FileSystem') { throw "Expected a filesystem directory: $OutDir" }
$outputPath = [IO.Path]::Combine($outputDirectory, $FileName)
if ([string]::Equals($outputPath, $workbook.FullName, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Output must be separate from the work book: $outputPath"
}
if ([IO.Directory]::Exists($outputPath)) { throw "Output is a directory: $outputPath" }
if (!$Force -and (Test-Path -LiteralPath $outputPath)) {
    throw "Output already exists: $outputPath. Use -Force to overwrite it."
}

$publishId = '.excel-publish-' + [guid]::NewGuid().ToString('N')
$updatedWorkbook = Join-Path $workbook.DirectoryName ($publishId + '.xlsm')
$pendingOutput = [IO.Path]::Combine($outputDirectory, $publishId + '.xlsm')
try {
    # Keep the work book intact until the imports and the release lock succeed.
    [IO.File]::Copy($workbook.FullName, $updatedWorkbook)
    & (Join-Path $PSScriptRoot 'Import-Vba.ps1') -Path $updatedWorkbook -SourceDir $sourceDir | Write-Verbose
    & (Join-Path $PSScriptRoot 'Add-CustomUI.ps1') -Path $updatedWorkbook -XmlPath $ribbonPath | Write-Verbose

    [void][IO.Directory]::CreateDirectory($outputDirectory)
    [IO.File]::Copy($updatedWorkbook, $pendingOutput)
    & (Join-Path $PSScriptRoot 'Lock-VbaProject.ps1') -Path $pendingOutput
    [IO.File]::Replace($updatedWorkbook, $workbook.FullName, [NullString]::Value)

    # Publish within the output directory so a partial copy is never visible.
    if ($Force -and [IO.File]::Exists($outputPath)) {
        [IO.File]::Replace($pendingOutput, $outputPath, [NullString]::Value)
    }
    else {
        [IO.File]::Move($pendingOutput, $outputPath)
    }
    $outputPath
}
finally {
    Remove-Item -LiteralPath $updatedWorkbook, $pendingOutput -Force -ErrorAction SilentlyContinue
}
~~~~~

### SKILL.md

<!-- excel-com-file SKILL.md 39aae1dcc2d2bba0fa4506bec6996861922fd787e4479de2ae9dded372a54044 -->
~~~~~markdown
---
name: excel-com
description: このリポジトリの Excel ブックを Windows 版 Excel と PowerShell COM で確認・編集する。シートの値・数式・書式、VBA と UserForm の作成・編集と入出力、カスタムリボンの埋め込みとボタンの動作確認、配布ブックの出力に使う。
---

# Excel COM

PowerShell から `Excel.Application` を COM で操作してブックを編集する。リボンだけは COM で定義できないため、閉じたブックの zip を直接書き換える。

コマンド例のパスは、リポジトリ直下から見た相対パス。Git Bash からは `/` 区切りのまま `powershell.exe -File` に渡してよい。

## 対象と作業手順

[AGENTS.md の「VBA とリボン UI の正本」](../../../AGENTS.md#vba-とリボン-ui-の正本)に従い、フォルダ名で対象を識別する。各スクリプトの `-Path` には、そのフォルダの `workbook.xlsm` を渡す。

```text
src/excel/<ブック名>/
├── workbook.xlsm
├── vba/
│   ├── standard/
│   ├── classes/
│   ├── documents/
│   └── forms/
└── ribbon/
    └── customUI14.xml
```

ブックを開く前に「マクロが動くことへの注意」と「作業前の確認」を読む。対象が特定できない場合は、作業ブックの場所を確認してから開く。

- **シートの確認・編集**は「COM 編集」に進む。確認だけなら `-Save` を付けない。
- **VBA の変更**は [VBA の編集ルール](references/vba.md)を読んでから `vba/` の正本を編集する。`Import-Vba.ps1` の `-SourceDir` に `vba/` を渡すと、その配下の対象ファイルをまとめて反映する。同期が不明なら、ブックのコードを別の作業先へ書き出し、反映対象すべての差分を確認する。変更したコンポーネントだけを反映する場合は、対象ファイルを種別ごとの配置を保って別の作業先へコピーし、そのルートを渡す。UserForm は付随する `.frx` もコピーする。
- **UserForm の作成・変更**は [UserForm の作成と編集](references/userform.md)を読む。VBE で確認できるよう、固定レイアウトはデザイナーに保存する。
- **リボンの変更**は `ribbon/customUI14.xml` を編集する。正本がない場合や同期が不明な場合は、ブック内の `customUI/customUI14.xml` を別の作業先に取り出して確認する。埋め込みは XML 全体を置き換えるため、既存の定義を残すか確認してから反映する。VBA も変更する場合は先に取り込む。ブックを閉じてから、XML を `Add-CustomUI.ps1` の `-XmlPath` に渡す。

編集後はブックを開き直し、「検証」で変更した箇所を確認する。

## 配布ブックの出力

閉じた `workbook.xlsm` をバックアップしてから、操作可能な Windows デスクトップで Windows PowerShell を使って実行する。VBA の取り込みと保護の確認には `AccessVBOM` が必要なので、「作業前の確認」に従う。実行中は Excel と VBE が表示されるため、先に一言伝え、処理が終わるまで表示されたウィンドウを操作しない。

```powershell
powershell.exe -NoProfile -File .agents/skills/excel-com/scripts/Publish-Workbook.ps1 -Path src/excel/sample-book/workbook.xlsm
```

一時コピーに `vba/` と `ribbon/customUI14.xml` を順に反映する。その内容を配布用に複製し、VBE の「プロジェクトを表示用にロックする」を設定する。パスワードは出力ごとに生成し、表示・記録しない。配布用コピーを保存して開き直し、ロック済みと確認できてから、作業ブックと配布先を更新する。取り込みやロックに失敗した場合は、元のブックと配布先を更新せずに終了する。

作業ブックは VBA を編集できる状態で更新する。配布後の修正も作業ブックと正本から行う。VBE のロックで制限するのは、Excel 上での VBA の閲覧・編集。ブック構成を保護するとシートの表示切り替えができなくなるため、配布時の保護は VBA プロジェクトに限る。

- `-FileName` は拡張子 `.xlsm` を含むファイル名を指定する。省略時はブックのフォルダ名に `.xlsm` を付ける。
- `-OutDir` は出力先のフォルダを指定する。省略時はリポジトリ直下に出力する。相対パスは実行時のカレントディレクトリから解決し、出力先フォルダがなければ作成する。
- 出力先に同名のファイルがある場合は、`-Force` を付けたときだけ上書きする。

## 手段の分担

| 対象 | 手段 | 理由 |
| --- | --- | --- |
| シート（値、数式、書式、シート構成） | COM | Excel が再計算し、保存形式も Excel が保つ |
| VBA モジュール | COM（`VBProject.VBComponents`、`scripts/Export-Vba.ps1` / `Import-Vba.ps1`） | `vbaProject.bin` の生成を Excel に任せられる |
| リボン（customUI） | zip を直接編集（`scripts/Add-CustomUI.ps1`） | COM にリボンを定義する API が無い |

順序は **COM の編集 → ブックを閉じる → customUI の埋め込み → COM で開き直して検証**。Excel で保存しても customUI は残るため、後から COM で再編集してよい。openpyxl で保存すると、リボンが読み込まれなくなるので混ぜない。openpyxl 3.1.5 の `keep_vba=True` で確かめた結果、`customUI14.xml` 自体は残るが `_rels/.rels` の関連付けが消え、Excel で開いても `onLoad` が動かなかった。VBA とセルの値は残る。混ぜてしまった場合は、`Add-CustomUI.ps1` をもう一度かければ関連付けが戻り、リボンも読み込まれる。

## マクロが動くことへの注意

シートの確認・編集と VBA の入出力では、`Use-ExcelWorkbook` の既定を使う。VBA とブックイベントを無効にし、外部ブックリンクを更新せずに開く。この設定は Excel 4.0 マクロやデータ接続の更新全般を抑止するものではない。

- VBA の動作検証に限り `-EnableMacros` を付ける。リボン押下のスクリプトはこの指定を付けて開く。実行前に `onLoad` などのコールバックも確認する。`Workbook_Open` などのイベントを検証する時は `-EnableEvents` も付ける。
- `$xl.Run` やボタンの押下で動かしてよいのは、この作業で自分が書いた、外部とやり取りしないマクロだけ。既存のマクロやボタンは、ユーザーの確認なしに動かさない。既存のブックには、リボンのボタンからサーバー上のデータを読み書きするものがある。ブックを保存せずに閉じても、サーバー側の変更は戻らない。
- COM で起動した Excel が sandbox の制限の内側で動くかは確かめていない。sandbox を、マクロに対する防御として数えない。

## 作業前の確認

1. **Excel のプロセス**: `Get-Process EXCEL`。動いているなら、ユーザーが開いているブックがある可能性がある。勝手に終了させず、閉じてよいか聞く。
2. **VBA を触るなら `AccessVBOM`**: `HKCU:\Software\Microsoft\Office\16.0\Excel\Security` の `AccessVBOM` が `1` でないと、`VBProject` へのアクセスが失敗する。これはセキュリティ設定なので、自分で変更せず、ユーザーに有効化を頼む。
3. **バックアップ**: 編集前のブックを scratchpad へコピーする。段階ごと（COM 編集後、customUI 埋め込み前など）にも取ると、どこで壊れたか切り分けられる。
4. **git 管理**: このリポジトリでは `*.xlsm` を既定で無視し、`src/excel/sample-book/workbook.xlsm` を例外として追跡する。未コミットの編集や追跡されないブックはバックアップから戻す。

## Excel COM の実行環境

- ユーザーが操作できる Windows デスクトップのセッションで実行する。実行環境を変えたときや、そこで COM の成功実績がないときは、作業ブックを開く前に、作業用の一時 `.xlsx` を `New-ExcelWorkbook` で作成する。`Use-ExcelWorkbook` でセルに値を書いて保存し、開き直して値を確認する。検証後は一時ブックを削除する。
- Codex の通常 sandbox 実行で Excel が制限時間内に COM 登録できず、正式な承認付き実行では同じ作成・保存・再読込が成功した例がある。その差を確認済みの環境では、同じ通常実行を繰り返さず、作業内容が承認済みであることを確認し、対象ブックとコマンドを限定して正式な実行権限を申請する。承認された実行経路で、マクロ・イベント・外部リンク更新を抑止する既定の手順を使う。
- 初めて `0x80080005` が出た環境では、対象ブックを使わない起動確認で実行経路の差を調べる。コードだけで sandbox が原因と断定しない。明示的な拒否や承認拒否があれば停止する。

## Excel COM が起動しないとき

- `New-Object -ComObject Excel.Application` で `0x80080005 (CO_E_SERVER_EXEC_FAILURE)` が出た場合は、対象ブックを開く前の起動失敗として扱う。HRESULT だけでは原因を特定できない。Excel の起動に成功してから `Workbooks.Open` や `VBProject` で失敗した場合は、それぞれの操作を切り分ける。
- 失敗した箇所が不明なときだけ、同じ実行環境で対象ブックを開かずに Excel COM の生成を1回試す。生成できたら `Quit()` と COM 解放を行う。既に生成単体で失敗を確認している場合は、同じ条件で繰り返さない。
- 起動前後の Excel プロセスを確認する。PID が増えただけでは、COM の生成成功や終了してよいプロセスであることは判断できない。ユーザーに普段の Windows セッションで Excel 単体が起動するか確認してもらう。レジストリ、DCOM、Office の設定は診断の一環として変更しない。
- COM の生成が失敗したままなら、Excel による編集・検証を止め、失敗した操作、HRESULT、対象ブックを開いたか、どこまで確認できたかを報告する。ブック内部を読み取り専用で調べた場合は、その結果と Excel での検証結果を分ける。

## COM 編集

バックアップや一時ファイルは、現在の権限で書き込める作業場所に置く。作業ディレクトリ外を使う場合は AGENTS.md の停止線に従う。PowerShell の実行ポリシーに阻まれたら、設定を変更せずに報告する。

作業用のスクリプトは scratchpad に置き、`powershell.exe -NoProfile -File <script> -Path <workbook>` で実行する。

Excel の起動、ブックを開く、後始末は `scripts/ExcelCom.psm1` の `Use-ExcelWorkbook` に任せる。ScriptBlock の中で例外が出ても、ブックを保存せずに閉じ、Excel に終了を求め、COM を解放する。どれか 1 つが失敗しても、残りは実行する。

```powershell
param([string]$Path)
$ErrorActionPreference = 'Stop'
Import-Module '<repo>\.agents\skills\excel-com\scripts\ExcelCom.psm1'

Use-ExcelWorkbook -Path $Path -Save -ScriptBlock {
    param($xl, $wb)
    $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = 'Item'
    $ws.Range('B2').Value2 = [double]3
    $ws.Range('D2').Formula = '=B2*C2'
}
```

新しいブックは `New-ExcelWorkbook -Path <file>` で作る。拡張子が `.xlsm` ならマクロ有効ブック、`.xlsx` なら通常のブックで保存する。既にファイルがあればエラーで止まる。作った後の編集は `Use-ExcelWorkbook` で行う。

- `-Save` は、ScriptBlock が例外なしで終わった時だけ保存する。検証のように読むだけなら付けない。
- `-Visible` はウィンドウを表示する。リボンを作らせる検証で使う。
- ScriptBlock が出力した値は、そのまま `Use-ExcelWorkbook` の戻り値になる。COM メソッドの戻り値（`Worksheets.Add()` など）も混ざるので、要らないものは `[void]` で捨てる。
- 終了後、この呼び出しで起動した Excel のプロセスが残っていれば、その PID を警告で出す。止める必要があれば、その PID だけを `Stop-Process -Id` で止める。名前（`-Name EXCEL`）では止めない。ユーザーの Excel まで止まる。読むだけの実行だった場合は、止めた後でブックがバックアップと同一か（`cmp`）を確かめる。

### 別ブックからシートをコピーするとき

- コピー元を同じ Excel インスタンスで読み取り専用で開き、リンク更新を抑止する。`Worksheet.Copy` で作業コピーにシートを追加し、コピー元を開いたままコピー先を保存する。コピー元を先に閉じると保存時に Excel が終了し、`0x800706BE` が返った例がある。`Use-ExcelWorkbook -Save` は ScriptBlock の後で保存するため、ScriptBlock 内でコピー元を閉じる場合は、ScriptBlock 内でコピー先を保存してから、`finally` でコピー元を閉じる。
- 保存後にコピー先を開き直し、定義名の名前・参照先・スコープと、入力規則・書式を元シートと照合する。`Worksheet.Copy` が成功しても、コピー先のブックでフォントが変わる場合がある。

### PowerShell + COM の癖

- **セルへの代入は `Range('B2')` 形式で、型を明示する。** `$ws.Cells.Item(r, c).Value2 = 3` は `System.Int32` を `System.String` にキャストできないという `InvalidCastException` で落ちることがある。`$ws.Range("B$r").Value2 = [double]$v` なら通る。
- **スクリプトは ASCII だけで書く。** `powershell.exe`（5.1）は BOM なしの `.ps1` を UTF-8 として読まない。日本語のシート名、モジュール名、VBA コードは、UTF-8 のデータファイルに分けて `[IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)` で読み込む。この方法で、日本語のモジュール名、コメント、文字列リテラル（全角と半角カナを含む）が化けずに入り、`$xl.Run` の戻り値も一致することを確かめてある。日本語のパスやラベルを `-File` の引数として渡すのは問題ない（Git Bash から `-TabLabel 'ツール'` のように渡して動くことを確かめてある）。
- **日本語の結果はファイルに書いて読む。** `powershell.exe` の標準出力は、Git Bash から見ると日本語が化ける。確認したい値は `[IO.File]::WriteAllLines($p, $lines, (New-Object Text.UTF8Encoding($false)))` で UTF-8 のファイルに書き出す。

## VBA の書き出しと取り込み

VBA をテキストで直す時や、ソースを Git で管理する時に使う。標準・クラス・文書モジュールと UserForm を扱う。シートとブックのモジュールを直接編集する場合は、`Worksheet.CodeName` または `Workbook.CodeName` で既存の `VBComponent` を取得し、その `CodeModule` に書く。

```powershell
powershell.exe -NoProfile -File .agents/skills/excel-com/scripts/Export-Vba.ps1 -Path <workbook> -OutDir <dir>
powershell.exe -NoProfile -File .agents/skills/excel-com/scripts/Import-Vba.ps1 -Path <workbook> -SourceDir <dir>
```

- **Export**: 標準モジュールは `.bas`、クラスと文書モジュールは `.cls`、UserForm は `.frm` と付随する `.frx` で書き出す。コードの無い文書モジュールは省く。今回書き出さなかったファイルが出力先に残っていれば警告する。
- **Import**: 同名・同種別の標準モジュール、クラス、UserForm は削除してから取り込む。文書モジュールは既存のコードを差し替える。対象ブックに同名の文書モジュールが無い場合は止まる。取り込み元に含まれないコンポーネントはブックに残る。最後に Excel で保存する。
- 文書モジュールの差し替えは、先頭のヘッダーと `Attribute` 行だけを除く。プロシージャの中の `Attribute` 行（`VB_Description` など）は扱わない。
- 既存のブックから書き出したコードには、接続先、ID、認証情報が入っていることがある。commit する前に中身を確かめる。

- `Export-Vba.ps1` は Excel が一時出力した CP932・CRLF のファイルを UTF-8（BOM なし）・LF に変換し、種別ごとのディレクトリに書き出す。
- `Import-Vba.ps1` はブックを開く前に、`.bas`、`.cls`、`.frm` が UTF-8（BOM なし）・LF で、CP932 に文字を失わず変換できることを確認する。VBE へは CP932・CRLF の一時ファイルを渡す。LF の `.cls` を VBE へ直接渡すと、標準モジュールとして認識された実測がある。

## リボン（customUI）

1. customUI XML を UTF-8（BOM なし）・LF で書く。名前空間は `http://schemas.microsoft.com/office/2009/07/customui`（Office 2010 以降用の `customUI14.xml`）。
2. ブックを閉じた状態で埋め込む。

   ```powershell
   powershell.exe -NoProfile -File .agents/skills/excel-com/scripts/Add-CustomUI.ps1 -Path <workbook> -XmlPath <customUI.xml>
   ```

   作業ブックと同じディレクトリの一時コピーに XML を埋め込む。`_rels/.rels` の関連付けは無い時だけ足す。書き出した XML と関連付けを読み戻して確認できた場合に、元のブックを置き換える。

3. ボタンの `onAction` などの callback は、VBA の標準モジュールに置く。シグネチャは `Public Sub Name(control As IRibbonControl)`、`onLoad` は `Public Sub Name(ribbon As IRibbonUI)`。
4. アイコンは組み込みの `imageMso`（例: `HappyFace`）を使える。

customUI の要素名や属性を間違えても、Excel は何も言わずにリボンを無視する。原因を見るには、Excel のオプション「詳細設定 → 全般 → アドインのユーザー インターフェイス エラーを表示する」が要る。これはユーザーの設定なので、必要ならユーザーに頼む。

## 検証

編集した後は、`Use-ExcelWorkbook`（`-Save` なし）でブックを開き直して読み戻す。書いたつもりの値ではなく、Excel が保存した結果を見る。動かしてよいマクロの範囲は「マクロが動くことへの注意」のとおり。

- **シート**: `Value2`、`Formula`、`Text`（表示形式が適用された文字列）を読む。数式は Excel が再計算した値で確かめる。
- **VBA**: 変更したコンポーネントの名前・種別・コードを読み戻す。実行してよいマクロは `$xl.Run('MacroName')` で戻り値を確かめる。ThisWorkbook に置いた関数は、`$xl.Run` で呼ぶと（ブック名で修飾しても）空が返る。`[System.__ComObject].InvokeMember('FuncName', [Reflection.BindingFlags]::InvokeMethod, $null, $wb, $null)` で Workbook のメンバーとして呼ぶと値が返る。
- **リボンとボタン**: `scripts/Invoke-RibbonButton.ps1` は表示名でタブを選び、そのタブの操作領域からボタンを探す。タブや操作領域が複数一致する場合や、同じタブ内でボタン名が重複する場合は、押す前にエラーで止まる。

  ```powershell
  powershell.exe -NoProfile -File .agents/skills/excel-com/scripts/Invoke-RibbonButton.ps1 -Path <workbook> -TabLabel <タブの表示名> -ButtonLabel <ボタンの表示名> -OutFile <result.json>
  ```

  結果は UTF-8 の JSON で書き出される。押せなかった時は終了コードが 1 で、`Error` に理由が入る。
  - `TabFound` と `ButtonFound` が `true` なら、リボンは読み込まれている。`Clicked` が `true` なら、ボタンを押す操作が成功した。
  - 押した後に出たダイアログ（MsgBox など）は、`Dialogs` に出た順に入る。`Names` には本文とボタン名が入る（例: `["OK", "処理が完了しました"]`）。MsgBox では本文もボタンも同じ種類の要素に見えるので、分けずに並べている。ダイアログは × ボタンと同じ操作で閉じる。
  - × で閉じられないダイアログ（Yes/No の MsgBox など）が出た時は、このスクリプトが起動した Excel だけを PID で止め、エラーにする。
  - 押す操作は、マクロの終了を待たずに戻る。ダイアログを出さずに長く動くマクロは、ブックを閉じる時に終わるまで待つ。終わらないマクロだと、スクリプトも終わらない。
  - 確かめた場面: 日本語のラベル、英字のラベル、ダイアログを出さないボタン、MsgBox を 2 回続けて出すボタン、Yes/No の MsgBox、押せない要素を指定した場合、30 秒動くボタン。
  - ブックは保存せずに閉じる。そのため、ボタンがシートに書いた値は、この結果からは見えない。シートへの効果は別に確かめる必要がある。
- **ウィンドウを表示すると、ユーザーの画面に Excel が数秒出る。** 先に一言伝えておく。リボンは表示したウィンドウで確かめている。非表示のまま開いた場合にリボンが作られるかは、確かめていない。

## このリポジトリとの関係

VBA を `.bas` / `.cls` ファイルとして保存・生成する時は、`AGENTS.md` の「`.bas` / `.cls` の保存・生成ルール」に従う。
~~~~~
