# T3 検証環境と初回コンソールの依存

- Date: 2026-11-01
- 出典: `dotnet --version`、`dotnet --list-sdks`、`dotnet --info`、Office登録情報、PIAのファイル情報・メタデータ・SHA-256、[開発ガイド](../development.md)、[Spike-0](../spike-0.md)。

## 対象と利用可否

T3は環境の静的確認と依存の選定までを対象とします。
既存設計のC#/.NET 10・Office Interop直接利用と、[T2の初期契約](../adr/0001-initial-lifecycle.md)を前提とします。
最初のコンソール検証は、既知のマクロなしブックの検証用コピーで起動・所有確認・保存・通常終了を確認する範囲です。
この範囲のT4準備へ進めると判断します。後続の[準備記録](spike-0-preparation.md)にビルド・非COM確認結果、[実機結果](spike-0-results.md)にユーザーの初回正常系とエージェントのCOM生成失敗を記載しています。

## 環境の確認結果

| 対象 | 観測・確認手段 | 判定と限界 |
| --- | --- | --- |
| .NET SDK | `dotnet --list-sdks`で10.0.402・10.0.403。リポジトリ内の`dotnet --version`は10.0.403 | SDKの選択を確認。[global.json](../../global.json)の既存規則を維持します |
| .NETのビット数 | `dotnet --info`のHostはx64、RIDはwin-x64 | x64ホストを確認。後続のビルド結果は準備記録を参照してください |
| SDK診断 | `dotnet --info`でService Control Managerへのアクセス拒否が発生 | 診断全体の成功とは扱いません。権限追加やworkloadの変更はしていません |
| Excel | ClickToRunのPlatformはx64、VersionToReportは16.0.20527.20088、ProductReleaseIdsはO365HomePremRetail。登録先EXCEL.EXEのProductVersionも一致 | ユーザーの了承に基づきMicrosoft 365 Excel 64-bitの登録情報を環境記録に使います。起動・接続による確認はしていません |
| Excel COM登録 | Excel.ApplicationのCLSIDとLocalServer32の登録を確認 | T3では登録の存在だけを確認。生成・終了の観測は実機結果を参照してください |

環境確認時のコマンドはリポジトリルートから実行しました。
登録情報は`HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration`と、ExcelのApp Paths・COM登録を読み取りました。
T3の環境確認ではExcelの起動・接続・終了、設定変更は行っていません。

## 設定の状態と必要になる工程

| 対象 | 確認結果 | 初回経路と後続の扱い |
| --- | --- | --- |
| ユーザー側のExcel Security設定 | T3の環境では`HKCU\Software\Microsoft\Office\16.0\Excel\Security`の読取りが拒否されました。T15の後の読取りでは、AccessVBOMは0、VBAWarningsは2でした。[T15の結果](spike-0-results.md#t15-vba経路の最小確認2026-11-01)を参照してください | マクロなしブックの通常操作ではVBIDEにアクセスしません。マクロ許可・VBIDEアクセスを初回準備の条件にしません |
| Excel Securityのポリシー | 調べたHKCU・HKLMの`Software\Policies\Microsoft\Office\16.0\Excel\Security`キーは見つかりませんでした | 他の管理経路を含む制限の不存在は証明していません。組織設定を変更・迂回しません |
| AutomationSecurity・イベント抑止・DisplayAlerts | T3では専用Excelを生成せず、インスタンスの値と効果を確認していません | [起動とコード実行の境界](../architecture.md#起動とコード実行の境界)に従うコードはT4で準備しています。効果はT5以降の該当検証で確認します |
| Protected Viewと終了安全確認 | T4の通常終了は実機結果を参照してください。混入・照会不能時の拒否は未確認です | 設定を緩和せず、T2の拒否条件を検証します |

VBIDE設定はVBA操作の前に確認します。マクロの実行や起動設定の効果はS10等で評価します。
読取り拒否を解除する操作やExcel画面での確認が必要になれば、対象と操作範囲を示して承認を得ます。
設定の存在や公式説明だけを、実効性の確認に代えません。

## 選定した依存と承認

2026-10-29に提示したT3プランへの「プランを実行して」というユーザー指示を、既存PIA直接参照の選定と、T4で次の2アセンブリを参照することへの承認として記録します。
T3では参照の追加やDLLのコピーを行っていません。T4では準備記録のとおり参照を追加しています。

| 参照対象 | AssemblyVersion / FileVersion | 用途 |
| --- | --- | --- |
| Microsoft.Office.Interop.Excel.dll | 15.0.0.0 / 15.0.4569.1506 | Application・Workbook・Workbooks・ProtectedViewWindows等の型 |
| OFFICE.DLL（assembly名: office） | 15.0.0.0 / 15.0.4613.1000 | MsoAutomationSecurity等のOffice共通型 |

型の存在はPEメタデータの読取りで確認しました。アセンブリのコードやCOMは実行していません。
既存ファイルを明示的に参照するため、新しいパッケージ配布元への依存を増やしません。
端末に依存する制約を受け入れ、T4では同じファイルを版とハッシュで照合します。
GACから.NET 10が自動解決するとは仮定しません。
ビルド時の型ライブラリ生成を伴うCOMReferenceは採用しません。[ResolveCOMReferenceの公式説明](https://learn.microsoft.com/en-us/visualstudio/msbuild/resolvecomreference-task?view=visualstudio)には、dotnet CLIからの利用制約があります。

ファイル・JSON・時間計測・診断は.NET標準機能を出発点とします。
MCP SDKの確認はR1の接続基盤整備へ前倒しします。状態は下記の追加確認を参照してください。VBIDE操作用の依存、ラッパー、汎用検証基盤は先行追加しません。
NuGetの[Microsoft.Office.Interop.Excel 16.0.18925.20022](https://www.nuget.org/packages/Microsoft.Office.Interop.Excel/16.0.18925.20022)は今回採用しません。
公開説明では第三者によるOfficeアセンブリの再梱包で、サポート対象外とされています。
このパッケージや別の版への変更は、配布元・利用条件・必要性を確認して改めて承認を得る対象です。

## 参照ファイルの識別と推移参照

パスは`%WINDIR%`を起点とします。DLLをリポジトリへ取り込んでいません。
AssemblyVersionだけでは同じファイルを特定できないため、SHA-256も記録します。

| ファイル | %WINDIR%からの相対パス | SHA-256 |
| --- | --- | --- |
| Excel PIA | assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll | 6F262BBA82EED8A8D69FAC44E491B99CCA2D4CD448166291CE2186833E730A85 |
| Office共通PIA | assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL | 5512CABD57B6E1FBD2B96C298D804A3795CD317F61E154AEDB335F6C119EAF62 |
| VBIDE PIA | assembly\GAC_MSIL\Microsoft.Vbe.Interop\15.0.0.0__71e9bce111e9429c\Microsoft.Vbe.Interop.dll | A8F8AB6C63C8D4471D158010F18CB24D4D2CCEA495A160CDCEF95A96183FFC6A |
| stdole | assembly\GAC\stdole\7.0.3300.0__b03f5f7f11d50a3a\stdole.dll | D6CDF961C6D2712FE1958815E51A30960D79FFF1E97788B7741627DBA972E8F7 |

Excel PIAのAssemblyRefはmscorlib 2.0.0.0、office 15.0.0.0、Microsoft.Vbe.Interop 15.0.0.0です。
Office共通PIAのAssemblyRefはmscorlib 2.0.0.0、stdole 7.0.3300.0です。
推移参照先のVBIDE PIAはAssemblyVersion 15.0.0.0、FileVersion 15.0.4561.1000でした。
stdoleはAssemblyVersion 7.0.3300.0、FileVersion 7.00.9466でした。
これらの参照情報は、初回経路で全アセンブリの明示参照が必要なことを証明しません。
VBIDE PIA・stdoleの明示参照は、初回コンソールのビルドには不要でした。後続で追加が必要なら根拠と影響を示して承認を得ます。
mscorlibの旧版を追加導入する判断ではありません。.NET 10での参照解決はビルドで確認します。

確認には`AssemblyName.GetAssemblyName`、`FileVersionInfo`、`Get-FileHash -Algorithm SHA256`を使いました。
型とAssemblyRefは`System.Reflection.PortableExecutable.PEReader`と`System.Reflection.Metadata`で読み取りました。
この確認方法は検証コードの依存として追加するものではありません。

## 未確認事項とT4への引継ぎ

- 上記2アセンブリの明示参照によるビルド結果は準備記録を参照してください。
- ビルドが追加依存や版の変更を要求する場合は、T3の選定を再評価します。新しい依存を無断で追加しません。
- Excel操作時の互換性、COM生成、所有確認・終了の確認範囲は実機結果を参照してください。古いmscorlib参照を含むPIAの存在やビルド成功だけで動作保証しません。
- 検証用資産とコードはT4で準備しています。Excelを扱う実行は対象と操作範囲の承認後に行います。
- T5の正常系と故障注入前の確認、R1の完了は、この環境調査では達成していません。

既存PIAの参照は初回検証に限定した、変更し直せる選択です。配布構成の確定ではありません。
Coreの参照は[ADR 0008](../adr/0008-embedded-office-pia.md)で同じファイルに決めました。
初回PIAの採用では公開契約・永続化・ライフサイクル設計を変更しないため、ADRは作成していません。

## MCP接続基盤の依存確認

2026-10-29に、R1で接続基盤を前倒しする実行プランが承認されました。[採用判断](../adr/0002-development-mcp-foundation.md)を参照してください。
固定対象は`ModelContextProtocol.Core` 2.3.0です。[公式リリース](https://github.com/modelcontextprotocol/csharp-sdk/releases/tag/v2.3.0)と、公式NuGetから取得したnupkgのnuspecを確認しました。
nuspecのrepository commitは`9a327829ef066c44a0d737c58bf3c7617d7c701c`でした。

### 取得経路と確認範囲

通常のエージェント実行からNuGetの版一覧を照会すると、名前解決が「そのようなホストは不明です」で失敗しました。
ユーザーの通常シェルでは指定版の存在を確認できたとの報告を得ました。
ユーザーはツールの昇格による照会・取得を明示承認しました。同じ照会は昇格した実行経路で成功し、指定版と推移依存のrestoreも成功しました。
実行経路による差を確認した結果です。名前解決を拒否した具体的な層・仕組みは未特定です。管理者権限や設定緩和が必要だったことを示しません。

復元にはサーバー直下の`NuGet.Config`を使い、配布元を公式NuGetだけに限定しました。
初回復元ではCLIホーム、APPDATA、global-packagesを`.xlwright/mcp-build/`配下に固定しました。HTTP・プラグインキャッシュの初回の保存先は未確認です。
継続利用するチェック入口ではHTTP・プラグインキャッシュも同じ作業領域へ明示的に固定します。証明書生成、ツールPATH追加、テレメトリは子プロセスの設定で抑止します。
依存取得以外のビルド・非COM検証は通常のエージェント実行で確認しています。実Excel操作の権限を変更したものではありません。

### 固定した依存

| パッケージ | 版 | nuspecのライセンス式 |
| --- | --- | --- |
| ModelContextProtocol.Core | 2.3.0 | Apache-2.0 |
| Microsoft.Extensions.AI.Abstractions | 10.11.0 | MIT |
| Microsoft.Extensions.Logging.Abstractions | 10.0.13 | MIT |
| Microsoft.Extensions.DependencyInjection.Abstractions | 10.0.13 | MIT |
| xunit.v3（テスト専用） | 4.0.2 | Apache-2.0 |

SDKは.NET 10向けアセットを使用します。直接依存と推移依存の版・content hashはCore、サーバー、2つのテストプロジェクトの`packages.lock.json`に記録しています。
xunit.v3の推移依存はApache-2.0またはMITです。製品のサーバーはこれらに依存しません。
テストはDLLを直接実行し、xUnitのIn-Process Runnerを使います。この経路はMicrosoft.Testing.Platformを起動しません。
推移依存のテレメトリ拡張は`dotnet test`やIDEの経路で登録されます。その経路の通信は未確認のため、テストは`npm run check`で実行します。
xunit.v3は2026-10-31にユーザーの承認を得て、公式NuGetから取得しました。
初回のネットワークrestore後は、通常のビルドでrestoreを行いません。依存変更時には版・利用条件・影響を確認します。
配布方式と第三者ライセンスの同梱方法は配布工程で確定します。ビルド成功を配布準備完了とは扱いません。

### 現在の検証範囲

開発用のExcel開始ツールとアプリ登録は、[ADR 0003](../adr/0003-inspector-centered-development.md)で削除しました。この節のアプリ接続とMCP経由の実行は、削除前の確認結果です。

接続基盤と開発用ツールのDebug・Releaseビルド、非COMチェックを確認しました。[結果記録](spike-0-results.md#mcp接続基盤の非com検証)を参照してください。
プロジェクト設定からのアプリ接続・診断と、固定コンソールによるMCP経由のExcel正常系2回を確認しました。受付ID形式の修正後の非COM17項目、アプリ再接続時の新しい起動識別子・ビルド一致、旧サーバー終了まで確認しました。開発ループの最小確認は完了しましたが、T5・R1全体の合格ではありません。
今回のstartingは根拠と対象を提示した承認後に復旧しました。今後も未完了記録の整理には同じ手動復旧の判断が必要です。

着手時のstartingのSHA-256は`9ACD9C6DD59C448899B0E2DF3930E5C9863EAAF982474A1397E1471F6C47EDE6`でした。
基準ブックのSHA-256は`E87C630F4B14250700A40359FD1788875D145AE2019974D8C5264D591BEBBDA2`でした。
事前のプロセス名照会でExcelは0件でした。これらの事前観測は、復旧根拠全体の確認や承認の代わりにはしません。
