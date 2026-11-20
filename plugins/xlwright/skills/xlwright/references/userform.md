# UserForm の作成と編集

## デザインと初期化の分担

フォームの固定レイアウトは、VBE でフォームを開いた時点で確認できるようにデザイナーへ保存する。タイトル、ラベル、固定のボタンやチェック欄、その配置・フォント・タブ順が該当する。[Microsoft Forms の Add メソッド](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms)

`UserForm_Initialize` には、フォームを開いた時に決まる状態の初期化を置く。入力データによって項目数や種類が変わる場合は、必要なコントロールを実行時に追加できる。[Initialize イベント](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/initialize-event)

## 作業コピーから正本へ反映

1. [作業前の確認](../SKILL.md#作業前の確認)に従い、対象ブックを閉じた状態でバックアップと作業コピーを作る。既存フォームは作業コピーから別の場所へ書き出し、正本と照合する。VBE の再書き出しでは、先頭付近の空行や識別子の大文字・小文字、`.frx` のバイト列が変わることがある（別のブックへ取り込むと、コードの先頭に空行が 1 行増えた実測がある）。差が出たらコードの意味を確認し、デザインは正本の `.frm` と `.frx` を別のコピーに取り込んで `Designer` のプロパティを比較する。残すべき変更を判断できない場合は、上書き前に確認する。
2. 作業コピーを `Use-ExcelWorkbook` で開き、既存の `VBComponent` を取得する。新規フォームは `VBComponents.Add(3)` で作る。フォームの `Caption`、`Width`、`Height` は `VBComponent.Properties.Item(...).Value`、コントロールは `VBComponent.Designer.Controls` から設定できる。`Properties` は、先に `VBComponent.Designer` を一度読むまで null になる（確かめた）。新しいコントロールは `Designer.Controls.Add('Forms.CommandButton.1', 'btnCancel')` のように、ProgID と名前で足す。足したコントロールは、足した時のフォームのフォントを受ける（`Name`、`Size`、`Bold`、`Charset` を確かめた。VBE のツールボックスから手で足しても同じ）。後でフォームの `Font` を変えても、既にあるコントロールのフォントは変わらない（確かめた）。フレーム内のコントロールはそのフレームの `Controls` に追加する。既存のコントロールを改名・削除する前に、イベントハンドラー名とコードからの参照を調べる。
3. 編集したコピーを保存し、`Export-Vba.ps1` で別の作業先に書き出す。対象フォームの `.frm` と `OleObjectBlob` が参照する同名の `.frx` を一組として扱う。書き出し前後でイベントコードを比較し、依頼に必要な初期化やイベントの差分だけを `.frm` に加える。イベントコードだけの変更では、既存のフォーム定義と `.frx` を使う。[VBA の編集ルール](vba.md)
4. 対象フォームの一組だけを `forms/` に置き、その親を `Import-Vba.ps1 -SourceDir` に渡して、編集前のブックの別コピーへ取り込む。次の検証が通った組を正本の `vba/forms/` と作業ブックへ反映する。

### PowerShell でフォントを設定できない場合

`Designer` から取得したコントロールに `.Font.Name` を代入すると、PowerShell 5.1 で `NullReferenceException` が出る。PowerShell から `.Font` を読んでも空になる。同じ代入を繰り返さず、`Designer` の `Font` を操作する一時 VBA を使う。一時 VBA は[マクロが動くことへの注意](../SKILL.md#マクロが動くことへの注意)に従い、`-EnableMacros` で実行し、保存・書き出しの前に取り除く。次の形で、フォントが `.frx` に保存されることを確かめてある。

```powershell
param($xl, $wb)
$components = $wb.VBProject.VBComponents
$tmp = $components.Add(1)
$tmp.CodeModule.AddFromString("Public Sub SetFont()`r`n    With ThisWorkbook.VBProject.VBComponents(""frmInfo"").Designer.Controls(""lblInfo"").Font`r`n        .Name = ""Meiryo UI""`r`n        .Size = 11`r`n    End With`r`nEnd Sub")
try { $xl.Run('SetFont') }
finally { $components.Remove($tmp) }
```

この例は 1 つのコントロールのフォントを変える。フォーム全体のフォントを変える時は、各コントロールに加えて、フォームの `Designer.Font` も同じ値にする。フォームの `Font` を変えないと、後で足したコントロールだけが元のフォームのフォントになり、他と違って見える。`Size` は入れた値に丸められることがある（11 を入れると 11.25 になった）。

## 取り込み後の検証

- **デザイン時**は、保存したブックをマクロ無効のまま開き直し、`VBComponent.Designer` から変更したプロパティを読み戻す。フォームとコントロールの名前、表示文言、位置・大きさ、タブ順、フレーム内の構成（`Parent.Name` とフレームの `Controls`）を変更内容に応じて確認する。コントロールは `foreach ($x in $designer.Controls)` でも、`Controls.Count` と `Controls.Item($i)`（0 始まり）でも列挙できる。フォントは PowerShell から読めないが、一時 VBA からなら `Designer` の `Font` を読める（確かめた）。読む関数は、一時フォルダに取ったブックの写しをマクロを有効にして開いてから足し、写しは保存しない。コントロールを足した時は、フォームと全コントロールの `Font` の `Name` と `Size` を読む。フォームの `Font` と違うフォントのコントロールがあれば、足したコントロールのフォントと並べて報告し、どちらに合わせるかをユーザーに確かめる。見出しだけ大きくするなど、意図して変えたものもあるので、自分では合わせない。VBE の見た目が依頼の中心なら、可能な環境ではデザイナー上の表示も確認する。
- **実行時**は、許可された範囲でフォームを表示せずに `Load` し、変更した初期化とイベントを確認する。たとえば、`Load frmInfo` の後にコントロールの値やフォントを返す関数を一時モジュールに置き、`$xl.Run` で呼んでから `Unload` する（`Show` はモーダルで止まるので使わない）。ボタンの Click は、一時モジュールの中で `frmInfo.btnCancel.Value = True` とすると起き、フォームが閉じたかは `VBA.UserForms.Count` で分かる。既存のボタンやマクロの実行は[マクロが動くことへの注意](../SKILL.md#マクロが動くことへの注意)に従う。デザインと実行時のどちらを確認したかを報告する。
- **見た目**（文字が切れる、コントロールが重なる、フォントが効いていない）は、数値からは分からない。`Save-WorkbookScreenshot.ps1 -UserForm <フォーム名> -EnableMacros` で、`UserForm_Initialize` の後のフォームを画像にし、Read で見る（[COM 編集の詳細](com.md)の「見た目を画像にする」）。フォームのコードが動くので、これも[マクロが動くことへの注意](../SKILL.md#マクロが動くことへの注意)に従う。
