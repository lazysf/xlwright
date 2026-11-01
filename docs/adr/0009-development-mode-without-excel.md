# 0009: 開発モードを製品ビルドに含め、Excelを起動しない

- Status: Accepted

2026-11-01にユーザーが承認しました。
`--development`は製品ビルドに残し、製品ツールと`dev_status`を公開します。
このモードの`session_start`は、設定の検証の後に`INVALID_REQUEST`、`detail: NOT_IMPLEMENTED`で止まり、所有記録にも触れません。
引数なしの起動は製品ツールだけを公開し、Excelを起動できます。

## 背景

Inspectorと`npm run check`は`--development`でReleaseのビルドを起動し、`dev_status`のハッシュでビルドを照合します。
製品ツールをInspectorで調査するには、このモードで公開する必要があります。
一方で、実Excelは固定コピーからだけ起動する規則です。

## 検討した代替案

製品ツールをそのまま公開する案は、Inspectorから固定コピー以外のビルドでExcelを起動できてしまいます。
開発モードを条件付きコンパイルで除く案は、Inspectorの検査と固定コピーの接続シナリオを別の構成へ作り直す必要があります。

## 影響

固定コピーの実Excelのシナリオは`--development`なしで起動します。
`Program.BuildHash`はsingle-file publishでは例外になります。呼出しは`dev_status`だけで、起動には影響しません。配布の方式を決めるときに直します。
