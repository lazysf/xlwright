# S14：計算結果の鮮度の固定検証

検証コードとVBAソースを準備しています。
2026-10-31 10:23 JSTのユーザー提示出力で、両check.ps1のDebug・Releaseビルドと非COMチェックの通過を確認しました。
Spike0自己チェック60件、MCP非COMチェック58件、両終了コード0です。
ケースA〜Cの観測・保存・後始末と計算鮮度の設計採用を確認し、T9を完了しました。
auto.xlsm・manual.xlsmと各sha256は、名前空間宣言を保持した補正候補の再検査後に配置済みです。
旧コピーは作成領域へ保全しています。
準備の根拠と限界は[T9の準備記録](../../../docs/notes/spike-0-results.md#t9のフィクスチャ準備2026-10-31)を参照してください。
資産がない状態の実行はExcel生成前に拒否します。
製品API・VBA反映・T10以降は実装しません。

## フィクスチャ作成の個別承認

実行するVBA全文は[S14Module.bas](S14Module.bas)です。
SetInput21はThisWorkbookのProbe!B2へ21を書きます。
SetManual31は専用Excelを手動計算へ変更して31を書きます。
SetAutomatic41は41を書いてから専用Excelを自動計算へ戻します。
ファイル操作・ネットワーク・外部プログラム起動・イベント・UDFは含みません。
作成・ビルド・各実験はそれぞれ別途承認が必要です。

1. ユーザーの通常PowerShell 5.1でExcel・旧実験コンソール・旧サーバー不在を確認します。
   所有記録completed・共有ガード不在・読取り排他も確認します。
2. エージェントがfixture-preparation配下のT9作成領域へ、作成用starter.xlsxとS14Module.basを用意します。
   starter.xlsxはExcelを起動せずに作成する補助資産です。
   ユーザーが新しく起動・識別した専用Excelでstarter.xlsxを開きます。
   既存ユーザー用Excelへ接続しません。
3. ブックを開いた状態で元の計算モードを記録します。
   設定済みのシートは順にProbe・Resultの2枚です。
   Probe!B2は数値1、Probe!C2は`=B2*2`、Result!B2は`=Probe!C2+1`です。
   補助資産のキャッシュを実Excelの計算確認には使いません。
4. S14Module.basをVBEから手動でインポートします。
   ThisWorkbook・各シートにコードを追加しません。
   VBAをコンパイルし、モジュールを書き出して承認ソースの本文と照合します。
   VBIDEへの自動アクセスや信頼設定変更は行いません。
5. 自動計算へ設定し、結果2・3を確認して作成領域のauto.xlsmへ保存します。
   手動計算へ設定し、同じ入力・結果のままmanual.xlsmへ別名保存します。
   CalculateBeforeSaveは変更せず、保存時の状態を記録します。
6. ブックを開いたまま元の計算モードへ戻して確認します。
   その変更は保存せず閉じ、専用Excelを通常終了します。
   未Open・Close後のCalculationは読みません。
7. ユーザーがプロセス終了・原本の読取り排他を確認します。
   原本は保持します。
   別コピーからworkbook.xmlのabsPath、文書プロパティのcreator・lastModifiedBy・Company・Managerを標準ZIP・XML操作で除去します。
   変更対象以外のパーツはVBAを含めバイト一致を確認します。
   homeパス・メール等の識別情報、外部リンク・接続・XLM・ActiveX・埋込みオブジェクトがないことも確認します。
8. シート・数式・初期キャッシュ・保存モードを確認したコピーをこのディレクトリへ新規固定します。
   各sha256は64桁のSHA-256と改行だけです。
   既存資産を上書きしません。

上記は作成手順です。
ユーザー原本は作成済みのため再作成しません。
名前空間保持の補正では既存の別コピーと配置済みファイルも証拠として保全します。
作成前の計算モードと復元値の具体的なユーザー確認は未取得です。
Excel自身の一時ファイル生成まで隔離する仕組みではありません。
予期しないコード・警告・アクセス拒否・内容不一致では停止します。

## ビルドと実験の承認

ユーザーがアプリを通常終了します。
通常PowerShell 5.1で旧サーバー・実験コンソール不在を確認してから、リポジトリルートで1本ずつ実行します。
両check.ps1はDebug・Releaseをビルドします。
エージェントはビルドしません。

```powershell
powershell.exe -NoLogo -NoProfile -File .\spikes\Spike0\check.ps1
$LASTEXITCODE
powershell.exe -NoLogo -NoProfile -File .\src\XlWright.McpServer\check.ps1
$LASTEXITCODE
dotnet --version
```

各終了コードと通過件数を保存します。
MCPの`-Restore`オプションは付けません。
実行拒否なら停止し、実行ポリシーを緩和しません。
再接続後はdev_status・起動記録・実DLLを照合します。
MCP開始ツールは使いません。

完成コード・固定ハッシュ・入力・副作用・終了方法の提示後、各ケースを個別承認します。
開始前と終了後のプロセス・読取り排他はユーザーの通常PowerShellを根拠とします。
以下は一覧です。
一括実行せず、承認された1本だけを標準入力を開いた対話コンソールで実行します。

```powershell
dotnet spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll --excel-s14-auto
$LASTEXITCODE
dotnet spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll --excel-s14-manual
$LASTEXITCODE
dotnet spikes/Spike0/bin/Release/net10.0-windows/Spike0.dll --excel-s14-mode-change
$LASTEXITCODE
```

| ケース | 設定・入力 | 仮説・期待観測 |
| --- | --- | --- |
| auto | 自動保存コピー。COMで11、SetInput21で21 | 明示計算なしの完了後に22・23、42・43。直後の旧値・新値は観測対象 |
| manual | 手動保存コピー。COMで11、SetInput21で21 | 直後と250ms後は前の結果。各変更後のApplication.Calculateと完了待ちで22・23、42・43。手動を維持 |
| mode-change | 自動保存コピー。SetManual31、明示計算、SetAutomatic41 | 手動化直後は2・3。明示計算後62・63。自動へ戻した後82・83 |

前ケースの証拠・終了・排他確認が完了するまで次へ進みません。
計算モードを測定前に統一しません。
Calculation・CalculationStateはOpen中だけ読みます。
250msの受動観測と50ms間隔の完了待ちは、同じSTAでメッセージ処理を行います。
操作全体の期限30秒に含めます。
開始・保存を含む終了は既存の各120秒です。
手動計算では受動観測でxlDoneを待ち続けません。
非同期外部更新や任意マクロの鮮度保証には広げません。

## 証拠・副作用・終了

既存の所有記録と共有ガードを使います。
T9は既存completedを要求します。
s14-auto.attempt等を新規作成して同じケースを拒否します。
後続ケースには前のT9のattempt・successの非空・一致が必要です。
T8 S4のsuccessは要求しません。

runへprobe.xlsm、diagnostics.jsonl、calculation-build.json、calculation-observations.json、calculation-saved.jsonを保持します。
各観測には値・型・数式・計算モード・読取り前後の計算状態・CalculateBeforeSave・経過時間を含めます。
複数のCOM読取りは原子的ではありません。
失敗前の観測は診断に残ります。
観測JSON全体は全測定成功後に作成します。
短い計算で状態遷移が見えなければ、待機の必要性を実証したとは扱いません。

書込み先はrun、今回の印・ガード、既存所有記録です。
新規依存・恒久設定・新規MCPツールはありません。
マクロは完全修飾した`'probe.xlsm'!S14Module.<手続き名>`で呼びます。
Open・COM編集・保存・Closeはイベント抑止下です。
マクロは元のイベント設定で呼び、警告抑止を追加しません。

正常時は保存、安全確認、Close、再確認、Quit、COM解放、対象終了確認、completed更新を通ります。
その後Excelなしで保存数式・値・モードと基準ハッシュを検査します。
successのflush・読戻しと自分のガード照合後だけ、そのガードを解除します。
保存時の再計算後の値を、変更直後の鮮度の証拠に使いません。
終了コード0とs14.completed、印・記録、ユーザー側終了・排他確認を照合します。

失敗・不一致は終了コード1、期限超過の結果不明は2です。
再送・自動代替・失敗後の設定復元や追加COM操作を行いません。
statusはCOMなしの状態照会です。
stop・EOF・Ctrl+Cは終了要求であり、COM停止の保証ではありません。
既存の安全確認付き終了要求経路を使います。
失敗証拠・印・ガードは保持します。
[初期の手動復旧](../../../docs/architecture.md#初期の手動復旧)に従い、旧プロセス終了確認後に人が対象と未保存内容を判断します。
復旧・記録整理は対象と証拠を示して別途承認します。
既存のT7・T8・復旧・空runは再実行・削除・移動・再利用しません。
約60秒の終了区間はT11・T12へ残します。
結果文書に個人パス・実PID・ログ全文を転記しません。
