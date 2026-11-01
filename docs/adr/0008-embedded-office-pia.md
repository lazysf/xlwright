# 0008: CoreのExcel型は既存PIAを埋め込んで参照する

- Status: Accepted

2026-11-01にユーザーが承認しました。
CoreはSpike0と同じGACのPIA 15.0.0.0を、`EmbedInteropTypes`で参照します。
リリース物の実行にPIAは不要で、ビルドにはPIAのある端末が必要です。

## 検討した代替案

NuGetの`Microsoft.Office.Interop.Excel` 16.0.18925.20022は、復元だけでビルドでき、CIを置けます。
一方で、[パッケージのページ](https://www.nuget.org/packages/Microsoft.Office.Interop.Excel/)は、再梱包のためライセンスがなく、サポート対象外と明記しています。
MITで公開する製品へこの由来の型を埋め込む問題のほうが、CIの利点より重いと判断しました。

`dynamic`の遅延バインドは依存を増やしません。全COM呼出しでコンパイル時の検査を失うため、採用しません。

## 影響

GitHubの標準ランナーと第三者の端末では、Coreをビルドできません。
`Formula2`のように版15にないメンバーは`dynamic`で呼びます。
CIや第三者のビルドが必要になったとき、および配布の着手時に、PIAの利用条件の確認を含めて見直します。
切替えの範囲はcsprojの参照です。
