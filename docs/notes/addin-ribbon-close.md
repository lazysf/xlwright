# アドインを閉じた後に残るリボンのタブ

- Date: 2026-11-20
- 出典: 一時フォルダの `.xlam` を、自動化で起動した Excel と、人が起動した Excel で開いて閉じた実測。自動化の側は、段階ごとに各ウィンドウのタブを UI Automation で列挙した。人の側は、アドインの VBA が書いたログと、人が画面で見た結果。Excel 16.0.20527.20196、Windows 11、Windows PowerShell 5.1.26100.9818。

アドインを閉じた後にリボンのタブが残って見える理由は、2 つあった。

- 前面になかったブックのウィンドウには、閉じる前のタブが残る。そのウィンドウを前面にすると消える。
- `Workbook_BeforeClose` で終了を取り消し、`Application.OnTime` で閉じ直すアドインでは、予約した時刻に閉じ直しが動かず、アドインが開いたままになる。タブはどのウィンドウにも出たままになる。

外部の実機調査で、試験用の `.xlam` を「再起動の後の 2 回目の終了」で閉じた時にタブが残った、と報告があった。報告のアドインは、`OnTime` で起動と終了を予約し、`Workbook_BeforeClose` で終了を取り消すことがある。配布用のアドインでは残らなかった。報告の手順の詳細は分からないが、2 つ目の理由で説明がつく。閉じ直しが遅れてアドインが開いたままになり、配布用のアドインはその場で閉じるので残らない。報告のアドインそのものでは確かめていない。

## 前面になかったウィンドウに残るタブ

利用側のコード（ブックのイベント、`OnTime`）を持たない `.xlam` で確かめた。

1. `.xlsm` に、MsgBox を出す callback の標準モジュールと、ボタン 1 つのカスタムタブ `AddinTab` を埋め込む。Excel で開き、`SaveAs(<path>, 55)` で `.xlam` にする（[export-addin.md](./export-addin.md) と同じ変換。ロックはしない）。
2. `Use-ExcelWorkbook -Visible -EnableMacros` で空の `host.xlsx` を開き、その Excel の中で次を順に行う。`.xlam` は `Workbooks.Open` で開き、`Close($false)` で閉じる。
3. 各段階の後に、表示中のウィンドウごとに `Get-RibbonTabs`（`Ribbon.psm1`）で `AddinTab` の有無を見る。2 回続けて同じ結果になるまで、最大 3 秒待つ。
4. 同じ手順を、新しい Excel で 2 回流した。閉じた後に前面を入れ替える段階を足して、もう 2 回流した。

| 段階 | 前面 | `host.xlsx` のタブ | `other.xlsx` のタブ |
| --- | --- | --- | --- |
| `.xlam` を開いた | host | あり | - |
| `other.xlsx` を開いて前面にした | other | あり | あり |
| `host.xlsx` を前面に戻した | host | あり | あり |
| `.xlam` を閉じた（1 回目） | host | なし | **あり** |
| `.xlam` を開き直した | host | あり | あり |
| `.xlam` を閉じた（2 回目） | host | なし | **あり** |
| `other.xlsx` を前面にした | other | なし | なし |
| 3 秒待った | other | なし | なし |
| `.xlam` を開き、`other.xlsx` を前面にして閉じた | other | **あり** | なし |
| `host.xlsx` を前面にした | host | なし | なし |

- 上の 6 行は 4 回とも、下の 4 行（前面を入れ替える段階）は 2 回とも、表と同じだった。
- 残ったタブは UI Automation の木だけでなく、ウィンドウの描画にも出ていた。`.xlam` を閉じた直後に各ウィンドウを `PrintWindow` で撮ると、`other.xlsx` のタブの列には `AddinTab` があり、`host.xlsx` には無かった。
- 前面にしないまま長く待った時に消えるかは見ていない。各段階で待ったのは、結果が 2 回続けて同じになるまでの 1 秒ほど。

## OnTime で閉じ直すアドイン

報告のアドインと同じく、`Workbook_BeforeClose` で 1 回目の終了を取り消し、`Application.OnTime Now + TimeSerial(0, 0, 3), "'<ブック名>'!DelayedClose"` で閉じ直しを予約する `.xlam` を作った。`DelayedClose` は、取り消さない印を立ててから `ThisWorkbook.Close False` を呼ぶ。各段階で、アクティブなブックの名前をログのファイルに書く。

### 自動化で起動した Excel

`Use-ExcelWorkbook -Visible -EnableMacros -EnableEvents` の中で試した。

| 予約したマクロ | 開き方 | 予約した時刻に動いたか |
| --- | --- | --- |
| `.xlam` の `DelayedClose`（`BeforeClose` から予約） | `Workbooks.Open` | 動かない（10 秒待った。3 回） |
| `.xlam` の `DelayedClose`（COM の `OnTime` で予約、モジュール名あり・なし） | `Workbooks.Open` | 動かない（6 秒待った） |
| `.xlam` の `Closer.DelayedClose` | `Use-ExcelWorkbook` で直接 | 動かない（8 秒待った） |
| `.xlsm` の `Closer.DelayedClose`（同じ VBA） | `Use-ExcelWorkbook` で直接 | 2 秒後に動いた |
| `.xlsm` の `Closer.DelayedClose`（同じ VBA） | `Workbooks.Open` | 動いた（8 秒以内） |

- 待つ間に COM を呼び続けた時も、呼ばずに `Start-Sleep` で待った時も、`.xlam` では動かなかった。
- `$xl.Run` で `.xlam` のマクロを直接呼ぶと動いた。マクロは有効だった。

### 人が起動した Excel

人が普段どおり Excel を起動し、空のブックを 2 つ開いてから、リボンに「終了」ボタンを足した同じ `.xlam` をダブルクリックで開いた。「終了」は `ThisWorkbook.Close False` を呼ぶ。

| 試験 | 操作 | ログ | 人が見たこと |
| --- | --- | --- | --- |
| A | Book1 で「終了」を押し、Book2 に切り替えた | 14:15:52 に取り消して予約。予約した 14:15:55 頃には何も無く、14:16:30〜33 に Excel を閉じようとした操作で取り消しが 3 回、その後に `DelayedClose` が動いて閉じた | タブはどちらのウィンドウにも出たまま |
| B | Book1 で「終了」を押し、Book1 のまま待った | 14:16:58 に取り消して予約。14:17:24〜27 に Excel を閉じようとした操作で取り消しが 3 回、その後に `DelayedClose` が動いて閉じた | タブは出たまま |

- 予約した 3 秒後には、どちらの試験でも閉じ直しが動かなかった。人が Excel を閉じようとした後に、やっと動いた。
- Excel を閉じようとすると、アドインの `BeforeClose` が取り消すので、Excel は 1 回では閉じなかった。

## このリポジトリでの扱い

- アドインを閉じた後にタブが消えたかを確かめる時は、先に `Workbooks.Item('<アドイン>.xlam')` などでアドインが閉じたかを確かめる。閉じていれば、確かめるウィンドウを前面にしてから見る。
- `.xlam` で `OnTime` に頼る処理は、予約した時刻に動かないことがある。試験は、`OnTime` を待たずに、予約するマクロを `Invoke-ExcelMacro` で直接呼んで確かめる。

## 未確認

- `.xlam` の `OnTime` が予約した時刻に動かない理由。公式の docs（`Application.OnTime`、2024-03-26 更新）は、Excel が Ready などのモードに無い時に待つことだけを書き、アドインには触れていない。
- `AddIns` に足して導入したアドインでも同じか。導入は人の Excel の設定を変えるので、試していない。
- 報告のアドインで、閉じ直しが遅れていたか。
- 人がウィンドウを並べて表示している時に、前面になかったウィンドウに残ったタブを押すとどうなるか。
