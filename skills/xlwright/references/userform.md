# UserForm の作成と編集

## デザインと初期化の分担

フォームの固定レイアウトは、VBE でフォームを開いた時点で確認できるようにデザイナーへ保存する。タイトル、ラベル、固定のボタンやチェック欄、その配置・フォント・タブ順が該当する。旧コントロールを隠して固定レイアウトを作り直す処理がある場合は、参照とイベントを確認したうえでデザイナー側の配置に移す。[Microsoft Forms の Add メソッド](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-microsoft-forms)

`UserForm_Initialize` には、シートの表示状態をチェック欄に反映する処理など、フォームを開いた時に決まる状態の初期化を置く。入力データによって項目数や種類が変わる場合は、必要なコントロールを実行時に追加できる。[Initialize イベント](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/initialize-event)

## 作業コピーから正本へ反映

1. [作業前の確認](../SKILL.md#作業前の確認)に従い、対象ブックを閉じた状態でバックアップと作業コピーを作る。既存フォームは作業コピーから別の場所へ書き出し、正本と照合する。VBE の再書き出しでは、先頭付近の空行や識別子の大文字・小文字、`.frx` のバイト列が変わることがある（別のブックへ取り込むと、コードの先頭に空行が 1 行増えた実測がある）。差が出たらコードの意味を確認し、デザインは正本の `.frm` と `.frx` を別のコピーに取り込んで `Designer` のプロパティを比較する。残すべき変更を判断できない場合は、上書き前に確認する。
2. 作業コピーを `Use-ExcelWorkbook` で開き、既存の `VBComponent` を取得する。新規フォームは `VBComponents.Add(3)` で作る。フォームの `Caption`、`Width`、`Height` は `VBComponent.Properties.Item(...).Value`、コントロールは `VBComponent.Designer.Controls` から設定できる。`Properties` は、先に `VBComponent.Designer` を一度読むまで null になる（確かめた）。新しいコントロールは `Designer.Controls.Add('Forms.CommandButton.1', 'btnCancel')` のように、ProgID と名前で足す。フレーム内のコントロールはそのフレームの `Controls` に追加する。既存のコントロールを改名・削除する前に、イベントハンドラー名とコードからの参照を調べる。`Caption` をシート名などのキーに使っている場合は、その対応も確かめる。
3. 編集したコピーを保存し、`Export-Vba.ps1` で別の作業先に書き出す。対象フォームの `.frm` と `OleObjectBlob` が参照する同名の `.frx` を一組として扱う。書き出し前後でイベントコードを比較し、依頼に必要な初期化やイベントの差分だけを `.frm` に加える。イベントコードだけの変更では、既存のフォーム定義と `.frx` を使う。[VBA の編集ルール](vba.md)
4. 対象フォームの一組だけを `forms/` に置き、その親を `Import-Vba.ps1 -SourceDir` に渡して、編集前のブックの別コピーへ取り込む。次の検証が通った組を正本の `vba/forms/` と作業ブックへ反映する。

### PowerShell でフォントを設定できない場合

`Designer` から取得したコントロールに `.Font.Name` を代入すると、PowerShell 5.1 と 7 で `NullReferenceException` が出る（5.1 で再現を確かめた）。PowerShell から `.Font` を読んでも空になる。同じ代入を繰り返さず、`Designer` の `Font` を操作する一時 VBA を使う。一時 VBA は[マクロが動くことへの注意](../SKILL.md#マクロが動くことへの注意)に従い、`-EnableMacros` で実行し、保存・書き出しの前に取り除く。次の形で、フォントが `.frx` に保存されることを確かめてある。

```powershell
param($xl, $wb)
$components = $wb.VBProject.VBComponents
$tmp = $components.Add(1)
$tmp.CodeModule.AddFromString("Public Sub SetFont()`r`n    With ThisWorkbook.VBProject.VBComponents(""frmInfo"").Designer.Controls(""lblInfo"").Font`r`n        .Name = ""Meiryo UI""`r`n        .Size = 11`r`n    End With`r`nEnd Sub")
try { $xl.Run('SetFont') }
finally { $components.Remove($tmp) }
```

## 取り込み後の検証

- **デザイン時**は、保存したブックをマクロ無効のまま開き直し、`VBComponent.Designer` から変更したプロパティを読み戻す。フォームとコントロールの名前、表示文言、位置・大きさ、タブ順、フレーム内の構成（`Parent.Name` とフレームの `Controls`）を変更内容に応じて確認する。コントロールは `foreach ($x in $designer.Controls)` でも、`Controls.Count` と `Controls.Item($i)`（0 始まり）でも列挙できる。フォントは PowerShell から読めないので、実行時に VBA で読む。VBE の見た目が依頼の中心なら、可能な環境ではデザイナー上の表示も確認する。
- **実行時**は、許可された範囲でフォームを表示せずに `Load` し、変更した初期化とイベントを確認する。たとえば、`Load frmInfo` の後にコントロールの値やフォントを返す関数を一時モジュールに置き、`$xl.Run` で呼んでから `Unload` する（`Show` はモーダルで止まるので使わない）。ボタンの Click は、一時モジュールの中で `frmInfo.btnCancel.Value = True` とすると起き、フォームが閉じたかは `VBA.UserForms.Count` で分かる。たとえばチェック欄がシートの表示状態に連動する場合は、作業コピーで表示状態を変えてからフォームを読み直す。既存のボタンやマクロの実行は[マクロが動くことへの注意](../SKILL.md#マクロが動くことへの注意)に従う。デザインと実行時のどちらを確認したかを報告する。
- **ブック全体**は、フォームだけを変更した場合、`.xlsm` の ZIP 内で `xl/vbaProject.bin` 以外の差分を確認する。Excel の保存で、無関係な XML や作業先の絶対パスが書き込まれた実測がある（`xl/workbook.xml` の `absPath` と、`docProps/core.xml` の更新日時）。同じ未署名ブックのコピーで、フォームの変更が VBA プロジェクト内で完結する場合は、元のブックのコピーに Excel が保存した `xl/vbaProject.bin` だけを移せる。その場合も他の ZIP エントリが元と一致することを調べ、完成候補を Excel で開き直してフォームを検証する。このスキルには `vbaProject.bin` を差し替えるスクリプトは無いので、`System.IO.Compression` で一時コピーのエントリを置き換え、エントリの順を保つ。
