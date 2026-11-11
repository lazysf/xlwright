# リボン（customUI）

コマンド例の `<skill>` は、SKILL.md に書いたこのスキルのフォルダの絶対パスに置き換える。

COM にはリボンを定義する API が無いため、閉じたブックの zip に `customUI/customUI14.xml` を直接埋め込む。

## 埋め込み

1. customUI XML を UTF-8（BOM なし）・LF で書く。名前空間は `http://schemas.microsoft.com/office/2009/07/customui`（Office 2010 以降用の `customUI14.xml`）。
2. VBA も変更する場合は先に取り込む。ブックを閉じた状態で埋め込む。

   ```powershell
   powershell.exe -NoProfile -File "<skill>/scripts/Add-CustomUI.ps1" -Path <workbook> -XmlPath <customUI.xml>
   ```

   作業ブックと同じディレクトリの一時コピーに XML を埋め込む。`_rels/.rels` の関連付けは無い時だけ足す。書き出した XML と関連付けを読み戻して確認できた場合に、元のブックを置き換える。埋め込みは XML 全体を置き換えるため、既存の定義を残すか確認してから反映する。正本がない場合や同期が不明な場合は、ブック内の `customUI/customUI14.xml` を別の作業先に取り出して確認する。
3. ボタンの `onAction` などの callback は、VBA の標準モジュールに置く。シグネチャは `Public Sub Name(control As IRibbonControl)`、`onLoad` は `Public Sub Name(ribbon As IRibbonUI)`。
   - callback には処理を書かず、引数の無いマクロ本体を呼ぶだけにする（例: `OnShowSum(control)` が `ShowSum` を呼ぶ）。マクロ本体を `$xl.Run('ShowSum')` で単独に試せ、シートのボタンなど別の入口からも使える。
4. アイコンは組み込みの `imageMso`（例: `HappyFace`）を使える。
5. 表示名は、ボタン押下で探せるように付ける。同じ表示名のタブを 2 つ作らない。1 つのタブの中で、ボタンの表示名を重複させない。どちらも `Invoke-RibbonButton.ps1` が押す前に `Ambiguous` で止まる。グループの表示名は、タブと同じでも、グループどうしで同じでもよい。

customUI の要素名や属性を間違えても、Excel は何も言わずにリボンを無視する。原因を見るには、Excel のオプション「詳細設定 → 全般 → アドインのユーザー インターフェイス エラーを表示する」が要る。これはユーザーの設定なので、必要ならユーザーに頼む。

## Excel 以外で保存しない

Excel で保存しても customUI は残るため、埋め込んだ後に COM で再編集してよい。openpyxl で保存すると、リボンが読み込まれなくなるので混ぜない。openpyxl 3.1.5 の `keep_vba=True` で確かめた結果、`customUI14.xml` 自体は残るが `_rels/.rels` の関連付けが消え、Excel で開いても `onLoad` が動かなかった。VBA とセルの値は残る。混ぜてしまった場合は、`Add-CustomUI.ps1` をもう一度かければ関連付けが戻り、リボンも読み込まれる。

## ボタンの押下で確かめる

`Invoke-RibbonButton.ps1` は、ブックを表示した Excel で開き（マクロ有効、ブックのイベントは無効）、表示名でタブを選び、そのタブの中からボタンを探して押す。ユーザーの画面に Excel が数秒出るので、先に一言伝える。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Invoke-RibbonButton.ps1" -Path <workbook> -TabLabel <タブの表示名> -ButtonLabel <ボタンの表示名> -OutFile <result.json>
```

結果は UTF-8 の JSON で書き出される。押せなかった時は終了コードが 1 で、`Error` に理由が入る。

- `TabFound` と `ButtonFound` が `true` なら、リボンは読み込まれている。`Clicked` が `true` なら、ボタンを押す操作が成功した。
- 同じ表示名のタブが複数ある場合や、同じタブの中でボタン名が重複する場合は、押す前に `Ambiguous` で止まる。
- 押した後に出たダイアログ（MsgBox など）は、`Dialogs` に出た順に入る。`Names` には本文とボタン名が入る（例: `["OK", "処理が完了しました"]`）。MsgBox では本文もボタンも同じ種類の要素に見えるので、分けずに並べている。ダイアログは × ボタンと同じ操作で閉じる。
- × で閉じられないダイアログ（Yes/No の MsgBox など）が出た時は、このスクリプトが起動した Excel だけを PID で止め、エラーにする。
- 押す操作は、マクロの終了を待たずに戻る。ダイアログを出さずに長く動くマクロは、ブックを閉じる時に終わるまで待つ。終わらないマクロは、上限で見張りが Excel を止める。
- ブックは保存せずに閉じる。そのため、ボタンがシートに書いた値は、この結果からは見えない。シートへの効果は別に確かめる。
- 確かめた場面: 日本語のラベル、英字のラベル、ダイアログを出さないボタン、MsgBox を 2 回続けて出すボタン、Yes/No の MsgBox、押せない要素を指定した場合、30 秒動くボタン、タブと同名のグループ、同名のグループが並ぶタブ、同名のタブ、同じタブ内の同名のボタン。
- 非表示のまま開いた Excel では、リボンは作られない（`onLoad` が動かない）。表示して開くと動く。リボンの確認には、必ずウィンドウを表示する。
