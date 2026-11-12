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

customUI の要素名や属性を間違えても、Excel は何も言わずにリボンを無視する。無視されたことは、`Test-RibbonLayout.ps1` でタブが `missing` になることで分かる（下の「見え方を XML と照らす」）。どの要素が誤りかを見るには、Excel のオプション「詳細設定 → 全般 → アドインのユーザー インターフェイス エラーを表示する」が要る。これはユーザーの設定なので、必要ならユーザーに頼む。

## Excel 以外で保存しない

Excel で保存しても customUI は残るため、埋め込んだ後に COM で再編集してよい。openpyxl で保存すると、リボンが読み込まれなくなるので混ぜない。openpyxl 3.1.5 の `keep_vba=True` で確かめた結果、`customUI14.xml` 自体は残るが `_rels/.rels` の関連付けが消え、Excel で開いても `onLoad` が動かなかった。VBA とセルの値は残る。混ぜてしまった場合は、`Add-CustomUI.ps1` をもう一度かければ関連付けが戻り、リボンも読み込まれる。

## 見え方を XML と照らす

`Test-RibbonLayout.ps1` は、ブックに埋め込まれた `customUI14.xml` を読み、表示した Excel（最大化、マクロ有効、ブックのイベントは無効）で XML の各タブを選び、UI Automation で読んだ見え方と照らす。ユーザーの画面に Excel が数秒出るので、先に一言伝える。

```powershell
powershell.exe -NoProfile -File "<skill>/scripts/Test-RibbonLayout.ps1" -Path <workbook> -OutFile <result.json> -ScreenshotDir <既存のフォルダ>
```

終了コードは、0 が一致、1 が食い違いあり、2 が確かめられなかった（customUI が無い、簡易表示のリボンなど。理由は `Error`）。

| 照らすもの | 対象 |
| --- | --- |
| タブ | 表示・非表示、`insertAfterMso` / `insertBeforeMso` の位置（間に挟まるカスタムタブは数えず、一番近い組み込みタブと比べる）、`startFromScratch="true"` で他の組み込みタブが消えていること |
| グループ | 表示名と順序。カスタムタブに入れた組み込みグループ（`idMso`）は位置だけ。組み込みタブ（`<tab idMso>`）では、足したグループの順序だけ |
| `button`、`toggleButton` | 表示名と順序、`size`（large / normal）、`enabled`、表示名と画像を持つ normal ボタンの `showLabel` |

`label` の無いグループとボタンも照らせる（UI Automation の名前は、グループが空白 1 字、ボタンが空。`screentip` があっても空のまま）。`visible="false"` のグループやボタンが出ていた時は、`Expected` が `hidden` の食い違いになる。

結果の JSON の見方:

- `Mismatches` に食い違いが `Path`（タブ/グループ/ボタン）、`Item`、`Expected`、`Actual` で入る。XML の誤りで Excel がリボン全体を無視した時は、タブが `missing` になる。
- `Dynamic` には callback（`getLabel`、`getVisible`、`getEnabled`、`getSize`、`getShowLabel`）で決まる値の、実際に出た値が入る。照合はしないので、設計どおりかを読んで確かめる。callback はブックのイベントを止めた状態で動くので、`Workbook_Open` で作る状態に頼る値は、普段と違うことがある。
- `Unsupported` には照らせなかったものが入る。menu、splitButton、gallery、box などを含むグループ（グループの有無だけ確かめ、中のボタンは照らさない）、callback で表示名が決まるタブ、`visible="false"` で隠した組み込みグループ（表示名が言語で変わり、探せない）、グループとボタンの `insertAfterMso` / `insertBeforeMso`（同じ理由で、組み込みの側を探せない）、`contextualTabs`。`separator` は UI Automation に出ないので数えない。
- `Collapsed` には、幅が足りずに 1 つのボタンにたたまれたグループが入る。中のボタンは照らさない。
- `Actual` には読んだタブ、グループ、ボタン（`Size`、`Enabled`、`IconOnly`）が入る。
- `-ScreenshotDir` を付けると、XML の順に `tab-1.png`、`tab-2.png` … を書き、`Screenshots` に場所が入る。アイコンの絵柄は UI Automation で読めないので、この画像を Read で見て確かめる。画像は画面ではなく Excel のウィンドウに描かせて撮る（`PrintWindow`）。手前に別のウィンドウを重ねた状態と、画面の拡大率が 100% 以外の時は確かめていない。画像を撮れなかったタブは、照合を止めずに `File` が `null`、`Error` に理由が入る。
- ツールチップ（`screentip`、`supertip`）は UI Automation で読めないので照らさない。

確かめた場面（Excel 16、Windows 11）: カスタムタブの `insertAfterMso`（同じ組み込みタブの後ろに 2 つ。XML で後に書いたタブの方が組み込みタブの近くに出た）、`getLabel` と `getVisible` を持ち隠れたボタン、large と normal、`showLabel="false"`、`enabled="false"`、`visible="false"`、`getEnabled` / `getVisible` / `getLabel`、toggleButton、カスタムタブに入れた組み込みグループ、組み込みタブに足したグループ（`insertAfterMso` 付き）、組み込みグループを隠す定義、`label` の無いグループとボタン、`startFromScratch`、menu を含むグループ、未知の属性でリボン全体が無視された場合。簡易表示のリボンの見分け（タブの中身の高さが 60 px 未満）は、実機で確かめていない。

## ボタンの押下で確かめる

`Invoke-RibbonButton.ps1` は、ブックを表示した Excel で開き（最大化、マクロ有効、ブックのイベントは無効）、表示名でタブを選び、そのタブの中からボタンを探して押す。ユーザーの画面に Excel が数秒出るので、先に一言伝える。

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
- 確かめた場面: 日本語のラベル、英字のラベル、ダイアログを出さないボタン、MsgBox を 2 回続けて出すボタン、Yes/No の MsgBox、押せない要素を指定した場合、30 秒動くボタン、タブと同名のグループ、タブと同名のシート、同名のグループが並ぶタブ、同名のタブ、同じタブ内の同名のボタン。
- ウィンドウが狭いと、Excel はグループを 1 つのボタンにたたみ、中のボタンは探せなくなる。そのため最大化して開く。最大化しても画面の幅が足りずにボタンが見つからない時は、`Error` にたたまれたグループの名前が付く。たたまれたかは、グループと同じ名前の要素 1 つがグループの高さいっぱいを占めることで見分けるので、グループと同名の menu 1 つだけのグループは、たたまれたと扱わない。どちらのスクリプトも、試験用の `-WindowWidth`（ポイント）を付けると、最大化せずにその幅で開く。`-Full` の試験は、これでたたまれる場面を作って確かめる。
- 非表示のまま開いた Excel では、リボンは作られない（`onLoad` が動かない）。表示して開くと動く。リボンの確認には、必ずウィンドウを表示する。
