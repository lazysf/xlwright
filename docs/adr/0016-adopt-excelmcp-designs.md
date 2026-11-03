# 0016: ExcelMcpのソースを読んで参考にし、xlwrightの方針で書く

- Status: Accepted

2026-11-03にユーザーが承認しました。
[ExcelMcp](https://github.com/sbroenne/mcp-server-excel)（sbroenne、.NET 10、MIT）には依存せず、併用もしません。
各段階で同じ対象のソースを読んで参考にし、xlwrightの要件と設計に合わせて、コードはxlwrightで書きます。
調べた事実は[調査記録](../notes/existing-excel-mcp-servers.md)にあります。

## 決定

ExcelMcpの形が、xlwrightの要件・設計・実機の確認の記録と食い違う時は、xlwrightに従います。
食い違う主な形は次のとおりです。

- 期限が切れても強制終了しない形と、前のサーバーが残したExcelを自動で片付けない形。[目標](../requirements.md#目標)の3と逆のためです。
- 記録と照合せずに、同じユーザーのExcelを列挙して接続先を選ぶ経路（`RunningExcelWindows.Find`と、それを呼ぶ接続の経路）。目標の4に反するためです。`XLMAIN`から`EXCEL7`を辿ってウィンドウから接続する手法（`Find`の中の辿り方と`Connect`）は、照合したPIDのウィンドウに限って、[片付け](../architecture.md#残った専用のexcelの片付け)の3の参考にします。
- PIDが取れない時に、Excelが生きているとみなす形（`ExcelBatch.IsExcelProcessAlive`）。設計の[所有の確定](../architecture.md#所有の確定)では、確かめられなければ止まるためです。

再試行の方針は、[設計](../architecture.md#期限と大きさの上限)のまま（`RETRYLATER`だけを合計10秒まで）とします。
ExcelMcpは、`RETRYLATER`を30秒、`SERVERCALL_REJECTED`を120秒まで再試行します（`OleMessageFilter.cs`）。REJECTEDの理由には、管理された端末で起動中に出るサインインのダイアログを挙げています。
xlwrightの方針の根拠（[T10](../notes/spike-0-results.md#t10の限定完了判定2026-10-31)）は制御したCOMでの評価で、実ExcelのRETRYLATERは未観測です。そのため、この差は設計の[構成](../architecture.md#構成)の未確認として扱います。

## 検討した代替案

- ライブラリとして依存する案。NuGetに出ているのは、.NETツールのMcpServerとCLIだけです。CoreとComInteropは、ライブラリとして出ていません。
- ソースを取り込んで直す案。方針の違う層も直せますが、MITの著作権表示が要ります。ユーザーは表示を置かないと決めました。本家は頻繁にリリースするため（[Releases](https://github.com/sbroenne/mcp-server-excel/releases)では、2026-11-02だけで2.6.8〜2.6.10の3回）、取り込んだ版との差も広がります。
- 2つのサーバーを併用する案。日常のExcelへ接続できる経路が残ります（目標の4）。同じ主ブックでは、排他と共有違反にぶつかります（目標の5）。
- ソースを読まず、文書だけを参考にする案。加速の効果が小さくなります。

## 影響

- ソースを横に開いて書き写すことはしません。MITの条件は、相当な部分を写す時に表示を残すことだからです。短い定型の書き方（Win32の呼出しの宣言など）は、この制約に当たりません。
- ExcelMcpのソースは、実機の確認の代わりにしません。
- 要件の目標は変えません。ExcelMcpは、人が開いて作業しているブックを扱うため、強制終了せずに人へ渡します。xlwrightは、非表示の専用のExcelで、Gitで管理された主ブックとテストの写しだけを扱います。そのため、目標の3の「保存していない変更は捨ててよい」が成り立ちます。
- 範囲の表も変えません。ExcelMcpには、[後続の範囲](../requirements.md#後続の範囲)の操作（ピボットテーブル、グラフ、Power Queryなど）の実装もあります。範囲へ加えるかは、人が判断します。
- PIAは[ADR 0008](0008-embedded-office-pia.md)のままとします。0から比べ直した結果です。
  - NuGetの`Microsoft.Office.Interop.Excel`は、[パッケージのページ](https://www.nuget.org/packages/Microsoft.Office.Interop.Excel/)が、再梱包のためライセンスが無く、サポート対象外と明記しています（2026-11-03に確認）。
  - `dynamic`は、すべての呼出しで型の検査を失います。
  - ExcelMcpも同じ名前空間の型を使うため、参考にした実装はPIA 15で書けます。版15に無いメンバーは、`dynamic`で呼びます。
- VBIDEは、設計のとおり遅延バインドで書きます。ExcelMcpの`VbaCommands`は`Microsoft.Vbe.Interop`の型で書かれていますが、VBIDEのPIAは加えません。
