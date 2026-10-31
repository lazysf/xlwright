# S9 STA・RETRYLATERの固定ケース

## 状態と範囲

T10は[ユーザー承認による限定評価を完了](../../../docs/notes/spike-0-results.md#t10の限定完了判定2026-10-31)しました。ユーザー実行のDebug・Releaseビルド、自己チェック71件、MCP非COM58件は通過しました。
両終了コードは0です。[確認範囲](../../../docs/notes/spike-0-results.md#t10のビルドexcelなし確認2026-10-31)を参照してください。
実機Aは0x800A03EC・終了コード1で失敗しました。RETRYLATER通知は未観測です。
再送とB・Cを停止しています。実Excel用の実行コマンドは現時点で使用しません。
承認済みの手動復旧は完了しました。[実測と復旧結果](../../../docs/notes/spike-0-results.md#t10ケースaセル編集状態でのsetter失敗2026-10-31)を参照してください。
実ExcelでのRETRYLATER発生、フィルターの有効性、操作期限と10秒上限の成立は未検証です。
T11・T12、製品API、製品サーバー内STA、PoC-1移行の合格を示しません。

基準は既存の`../lifecycle.xlsx`です。原本をExcelで開きません。
実行時に既存の`.xlwright/spike0/runs/<run-id>/`へ新規コピーを作ります。
VBA・VBE・ActiveX・新しい依存・恒久設定は使用しません。

## 実装と証拠

実行コードは`spikes/Spike0/S9Probe.cs`と`ExcelS9Probe.cs`です。
既存のProgram、ExcelScenario、ExecutionGate、所有記録、共有ガードを通ります。
フィルターはsetterの区間だけ同じSTAへ登録し、元へ復元します。
元のsetterをC#から投入するのは1回です。アプリケーションによる再送はしません。

再試行は`RetryRejectedCall`の`dwRejectType=2`だけです。
間隔は100msです。残り操作期限または再試行予算が100ms未満なら中止します。
拒否を含む呼出しの経過時間を保守的に操作全体の10秒予算へ計上します。
許可時刻と実復帰は別に記録します。予約後の厳密な送出時刻の保証ではありません。
MessagePendingは標準待機処理を返し、受理済み呼出しを取り消しません。

runにはコピー、diagnostics.jsonl、s9-returned.json、期限到達時はs9-timeout.jsonを保持します。
S9証拠にはケース、ランタイム、実DLLハッシュ、状態、投入回数、STA情報、拒否種別、tick値、単調時計、待機判断、残予算を含みます。
最終HRESULTと失敗段階はdiagnostics.jsonlへ記録します。
終了後の成功経路では既存の保存XML検査でA2=1を照合します。
フィルター内ではファイル・標準出力・COMを呼びません。

## ビルドとExcelなしチェック

ユーザーがアプリを通常終了し、旧サーバー・実験コンソール不在を通常PowerShellで確認します。
Windows標準PowerShell 5.1で、リポジトリルートから1本ずつ実行します。
実行拒否・失敗なら停止します。ポリシーを緩和しません。

```powershell
powershell.exe -NoLogo -NoProfile -File .\spikes\Spike0\check.ps1
$LASTEXITCODE
powershell.exe -NoLogo -NoProfile -File .\src\XlWright.McpServer\check.ps1
$LASTEXITCODE
dotnet --version
```

両チェックはDebug・Releaseをビルドします。MCP側に-Restoreは付けません。
エージェントはビルドも自己チェックも実行しません。
S9ChecksはWinForms投稿処理の配送・同じSTA・フィルター登録復元をExcelなしで検査します。
偽時計では拒否種別、10秒境界、残り100ms未満、呼出し間の予算共有を検査します。
受理済み操作の遅延成功・例外では後続の書込み・保存・Close・記録更新と完了受理を拒否します。
これは合成確認です。実Excelの拒否・遅延復帰の観測には代用しません。

## 制御した実COM試験（Excelなし）

`--s9-com-check`は別プロセスの固定受信フィクスチャとの実COM通信を検査します。
追加後のDebug・Releaseビルドと自己チェック71件・MCP非COM58件、制御COM試験7ケースはユーザー実行で通過しました。
制御COM試験の終了コードは0です。[実測と限界](../../../docs/notes/spike-0-results.md#t10の制御com試験2026-10-31)を参照してください。
この固定試験は再実行しません。以下のコマンドは実測時の再現情報です。
偽時計の合成検査や、実ExcelのRETRYLATER観測とは区別します。

実装は`spikes/Spike0/S9ComCheck.cs`と`S9ComReceiver.cs`です。
同じビルド済みDLLからケースごとに受信プロセスを1つ起動します。
受信側はStandardOleMarshalObjectと標準IStreamを使い、メモリ内のCommitだけを実装します。
他のストリーム操作は未対応として例外にします。
CoMarshalInterfaceとCoUnmarshalInterfaceによるローカルプロセス間の標準マーシャリングを使います。
オブジェクト参照は自分で起動した受信プロセスの標準出力からだけ受け取り、実行識別子・ケース・PID・DLLハッシュを照合します。
外部ファイルやユーザー引数から参照を取り込みません。参照の生データは証拠へ保存しません。
COM登録、レジストリ変更、依存追加、Excel起動はありません。
受信側フィルターはIStream.Commitだけを拒否し、準備や参照管理の呼出しは拒否しません。
呼出し側は既存のRetryFilterとExecutionGateを使います。

| ケース | 固定条件 | 期待観測 |
| --- | --- | --- |
| accept | 直ちに受理、操作30秒 | 投入1回・受信STAで本文1回・成功・後続マーカー実行 |
| retry | 最初の2回をRETRYLATER、その後受理、操作30秒 | 通知type=2・100ms再試行・投入1回・本文1回・成功 |
| reject | SERVERCALL_REJECTED、操作30秒 | 通知type=1・再試行なし・本文0回 |
| limit | RETRYLATERを継続、操作30秒・予算10秒 | 予算で打切り・本文0回・後続なし |
| deadline | RETRYLATERを継続、操作2秒・予算10秒 | 操作期限で打切り・isolated・本文0回・後続なし |
| failure | 受理後に固定E_FAIL、操作30秒 | 本文1回・再送なし・後続なし |
| late | 受理後に500ms待って成功、操作100ms | 保留中のisolated・busyを観測し、復帰後もisolated・後続なし |

後続マーカーは管理コードだけです。実Excelの保存・Close等は実行しません。
失敗・期限切れを期待するケースのPASSは、その拒否・抑止の評価成功を意味します。
元のCOM操作が成功したという意味ではありません。
時間境界の厳密な送出時刻を保証する試験ではありません。
観測不足・判定不一致なら次のケースへ進みません。

アプリ通常終了と旧サーバー・実験コンソール不在確認後、通常PowerShell 5.1で両check.ps1を実行します。
両方が終了コード0のときだけ、同じリポジトリルートから次を1回実行します。
エージェントはビルド・実COM試験を実行しません。

```powershell
dotnet spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll --s9-com-check
$LASTEXITCODE
```

受信側の`--s9-com-receiver`は専用の内部入口です。手動では起動しません。
各ケースで参照解放後にパイプで通常終了を要求し、受信結果と対象プロセス終了を確認します。
失敗時も入力パイプを閉じて通常終了を要求します。COMの取消しやプロセスの強制終了はしません。
準備応答10秒、クライアントSTAの観測45秒、受信結果・終了確認はそれぞれ10秒を上限とします。
観測期限は処理停止を保証しません。終了未確認なら残存可能性を残して停止します。

証拠は`.xlwright/spike0-com-checks/<run-id>/`へ新規作成します。
各ケースのprocess.json、client.json、receiver.json、stderr.log、assessment.jsonを保存します。
例外時はunknown.jsonを保存します。途中失敗では全種類が揃うとは限りません。
clientにはHRESULT、投入回数、期限観測、フィルターの時刻・判断列を記録します。
receiverには受信・拒否・本文実行回数とフィルター復元を記録します。
全ケースの評価と終了確認後だけcompleted.jsonを作り、今回の専用active.jsonを除去します。
失敗時はactive.jsonを保持して再実行を拒否します。削除して続行しません。
既存のExcel用所有記録・ガード・実行印・復旧証拠は読取り・更新とも行いません。

この試験が通過しても実ExcelのRETRYLATERは未観測のままです。T10の完了は上記の限定判定によるもので、実Excel側の成立確認には広げません。

## 実Excelの承認と実行

準備の承認は実Excel実行の承認を含みません。
完成コード・入力・副作用・終了方法の確認と、各ケースの個別承認が必要です。
開始前と終了後に、ユーザーの通常PowerShellで対象プロセス存否、所有記録、ガード、読取り排他取得と解放を確認します。
過去の不在・排他の証拠は流用しません。開始ツールは使いません。

既存所有記録completedと共有ガード不在を要求します。
各ケースのs9-release.attempt等はCreateNewで作り、再送を拒否します。
B・CにはAのattemptとsuccessの一致が必要です。CにはBのattemptも必要です。
C実行前にはBの評価と別承認の復旧が必要です。この人の承認をCLIが認証するものではありません。

各ケースは別の専用Excelと新規コピーを使用します。
Captionの実行識別子とコピーを確認してください。既存ユーザー用Excelへ接続しません。
初期値A2=0と対象Rangeを先に取得し、イベントを抑止します。
コンソールの準備合図後、コピーのLifecycle!D2を選び、F2で編集状態にします。
文字を入力せず、コンソールへ戻ってgoと入力します。標準入力は開いたままにします。
その後、C#から`target.Value2 = 1d`を1回投入します。

| ケース | 仮説・設定・期待観測 | 終了 |
| --- | --- | --- |
| A release | 操作30秒・予算10秒。最初のRETRYLATERから約2秒後の解除合図で、対象Excelへ戻りEsc。拒否後に同じ元呼出しが成功し、読戻し1 | 保存・安全確認・Close・Quit・対象終了・completed・保存検査後だけ終了コード0とsuccess |
| B limit | 操作30秒・予算10秒。編集を維持。操作期限より先に予算で打切り | 終了コード1を想定。保存・Close・Quitなし。記録・ガードを保持 |
| C deadline | 固定操作期限2秒・予算10秒。編集を維持。操作期限でisolated、後続を抑止 | 終了コード2を想定。記録・ガードを保持 |

```powershell
dotnet spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll --excel-s9-release
$LASTEXITCODE
```

```powershell
dotnet spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll --excel-s9-limit
$LASTEXITCODE
```

```powershell
dotnet spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll --excel-s9-deadline
$LASTEXITCODE
```

まとめて実行しません。順序はN → A → B → 個別承認した復旧 → C → 個別承認した復旧です。
セル編集中にRETRYLATERが発生することは仮説です。
未観測・別拒否・保留なら停止し、別のUIやVBAへ切り替えません。
期限直前の100ms未満では再試行せず、Cでその拒否が返った場合はCOMなしで実期限到達を観測します。
期限切れはExcel側の停止を証明しません。コンソール終了後のCOM復帰回収も保証しません。

## 停止・復旧・判定

失敗・結果不明・証拠不足では再送せず、以後のCOM・保存を止めます。
B・Cで予想した失敗でも通常成功印は作りません。ガードを保持します。
stop・EOF・Ctrl+Cは停止要求で、COM停止の保証ではありません。
停止要求の場合だけ既存の安全確認付き終了経路を使う可能性があります。
故障注入としてのEOF・終了実験は実施しません。

旧コンソールと旧サーバーの終了確認後、人が対象Excelの未保存内容と保存・破棄を判断します。
対象Excelの終了を確認し、証拠を保全してから、個別承認した記録・ガードだけを整理します。
所有不明のプロセスは終了しません。強制終了・自動復旧はありません。
過去の実行印・成功印・空run・復旧証拠は削除・移動・再利用しません。

判定は合成チェック、実機観測、未検証の仮説を分けます。
期待した打切りでも、時刻・拒否種別・投入1回・後続処理なしを証拠で照合します。
終了後は基準ハッシュとGit差分を確認します。
結果文書に個人パス・実PID・ログ全文を転記しません。
