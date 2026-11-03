# Excelを扱う既存のMCPサーバーの調査

- Date: 2026-11-03
- 出典: [sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)（2026-11-03に読んだmain。その時点の最新の公開版は2.6.10。文書は[ADR-001](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/ADR-001-TESTING-STRATEGY.md)、[ADR-004](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/ADR-004-GENERATED-COMMAND-CONTRACTS.md)、[ADR-005](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/ADR-005-SESSION-AND-FAILURE-OWNERSHIP.md)、[ARCHITECTURE.md](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/ARCHITECTURE.md)、[COM-API-BEHAVIOR-FINDINGS.md](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/COM-API-BEHAVIOR-FINDINGS.md)、[PIA-COVERAGE.md](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/PIA-COVERAGE.md)、[RUN-VBA-MACROS.md](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/guides/RUN-VBA-MACROS.md)、[AUTOMATION-ADVANCED.md](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/features/AUTOMATION-ADVANCED.md)。主に読んだソースは`src/ExcelMcp.ComInterop/`の`OleMessageFilter.cs`、`Session/`の`ExcelBatch.cs`・`ExcelDialogProbe.cs`・`ExcelShutdownService.cs`・`OwnedProcessGuard.cs`・`ProcessTerminationPolicy.cs`・`RunningExcelWindows.cs`、`src/ExcelMcp.Core/Commands/Vba/`、`ExcelMcp.ComInterop.csproj`、`ExcelMcp.Core.csproj`。ほかに[Releases](https://github.com/sbroenne/mcp-server-excel/releases)）、[xlide_mcp](https://github.com/WilliamSmithEdward/xlide_mcp)、[ThepExcelMCP](https://github.com/ThepExcel/ThepExcelMCP)、[prizmPrograms/excel-mcp-server](https://github.com/prizmPrograms/excel-mcp-server)、[Excellm](https://github.com/mroshdy91/Excellm)、[haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server)、[Agent Mode in Excel](https://techcommunity.microsoft.com/blog/excelblog/agent-mode-in-excel-is-now-generally-available-on-desktop/4457408)

xlwrightを0から作る前に、[要件](../requirements.md)を満たす既存のMCPサーバーがあるかを調べました。
各リポジトリはREADMEと文書を読みました。ソースは、ExcelMcpの出典に挙げたファイルを主に読みました。
xlwrightの要件の中心である、[目標](../requirements.md#目標)の2〜5の仕組みと、ファイルを正本としたVBAの反映を満たすものはありませんでした。
ExcelMcpの扱いは[ADR 0016](../adr/0016-adopt-excelmcp-designs.md)で決めました。

## 候補

| MCPサーバー | Excelとのつなぎ方 | VBA | 実装と規模 |
| --- | --- | --- | --- |
| ExcelMcp（sbroenne） | 自分でExcelを起動します。動いているExcelへの接続も、明示すればできます | 一覧、閲覧、モジュールの置換、`.bas`の取り込み、期限付きの実行 | .NET 10、MIT。64ツール・437操作。2026-11-02に2.6.10を公開 |
| xlide_mcp | 読み書きは、Officeを使わずにファイルを扱います。実行は、自分で起動したOfficeで行います | テストの実行、コンパイルの確認、`.bas`と`.cls`の書き出しと取り込み | Python、MIT、スター11 |
| ThepExcelMCP | 動いているExcelに接続します。最初のExcelに無いブックは、ROTから探します | 一覧、読み書き、削除、実行。既定で無効 | Python、MIT、スター16 |
| prizmPrograms/excel-mcp-server | 動いているExcelに接続します | 追加、編集、実行。エラーを捕まえる一時のモジュールを作ります | TypeScriptとPowerShell、MIT、スター5 |
| Excellm | 動いているExcelに接続します | 実行。既定で無効 | Python、MIT、スター1 |
| haris-musa/excel-mcp-server | Excelを使わず、openpyxlでファイルを扱います | 読み書きだけで、実行しません | Python、MIT、スター4.3k |

スターの数は2026-11-03に見た値です。
MicrosoftのExcel用のMCPサーバーは、公式の情報に見つかりませんでした。CopilotのAgent Modeは、MCPサーバーではありません。

## ExcelMcpの作り

ExcelMcpは、候補の中で最も活発で、Excelの寿命と失敗の扱いを文書にしています。

- 構成: `src/`は、ComInterop、Core、Service、McpServer、CLI、生成器に分かれます。NuGetに出ているのは、.NETツールのMcpServerとCLIだけです。
- PIA: NuGetの`Microsoft.Office.Interop.Excel`を、実行時に含めず、型を埋め込んで参照します。VBEとOffice共通の型は、`ExcelDna.Interop`から取ります（`PIA-COVERAGE.md`）。
- 起動と接続: 自分のExcelは`Activator.CreateInstance`で起動します。動いているExcelへは、`RunningExcelWindows`で接続します。`Find`は同じユーザーのExcelを列挙し、`XLMAIN`の子の`EXCEL7`を辿って接続先を探します。`EXCEL7`の無いExcelは、結果に入りません。`Connect`は、PIDと生成時刻を照合してから接続します。
- 終了: PIDと生成時刻で照合します。終了の直前にハンドルを開き直し、もう一度照合してから`TerminateProcess`を呼ぶ関数があります。ただし、製品の経路からは呼ばれていません。呼出し元が無いことは、その時点の`src/`全体を検索して確かめました。
- 終了の待ち: `ProcessTerminationPolicy.cs`は、Quitの後の自然な終了を75秒、終了させた後の確認を5秒とします。75秒の待ちは起動の失敗の後始末で使い、過ぎても終了させません。
- メッセージフィルター: 専用のSTAに登録し、`RETRYLATER`を30秒、`SERVERCALL_REJECTED`を120秒まで再試行します。REJECTEDの理由には、管理された端末で起動中に出るサインインのダイアログを挙げています。
- ダイアログ: 所有するPIDの`XLMAIN`が無効の時に、表示されて有効なウィンドウの所有者を辿り、`XLMAIN`に届くものを見つけます。クラス名の`#32770`は使いません。ダイアログの文字は、500ミリ秒で区切って読みます。
- PIDが取れない時は、Excelが生きているとみなします（`ExcelBatch.IsExcelProcessAlive`）。
- テスト: 振る舞いを持つ境界で試します（ADR-001）。Excelに依存しない処理はExcelなしで、ブックの操作は実Excelで試します。
- 契約の生成: Coreの注釈から、Serviceの振り分け、CLI、MCPのツールを生成します（ADR-004）。xlwrightはツールが33個なので、この仕組みは取り入れません。
- `COM-API-BEHAVIOR-FINDINGS.md`は、ほぼPower QueryとData Modelの知見です。xlwrightの範囲とは、ほとんど重なりません。

## ExcelMcpとxlwrightの要件の差

次の点は、xlwrightの[設計](../architecture.md)と同じ考え方です。

- 自分で起動したExcelを、PIDと起動時刻で識別します。プロセス名でExcelを終了させません。
- COMの呼出しを専用のSTAで行い、OLEのメッセージフィルターを登録します。
- 期限切れや確かめられない終了を結果不明として扱い、前の編集を再送しません。
- 開く時は、既定でマクロを無効にします。実行するには、開き直しの承認が要ります。

次の点が、xlwrightの要件と食い違います。

- 期限が切れても、Excelを強制終了しません。Excelを人に見せ、判断を人に渡します。目標の3は、上限を過ぎた専用のExcelを終了させ、AIエージェントが作業を続けることを求めています。
- 前のサーバーが残したExcelは、自動では片付けません。目標の3は、次の開始で片付けることを求めています。
- 記録と照合せずに、同じユーザーのExcelを列挙して接続先を選ぶ経路（`RunningExcelWindows.Find`と、それを呼ぶ接続の経路）を持ちます。目標の1のためにツールをすべて事前に許可すると、AIエージェントが日常のExcelに接続できます。目標の4は、日常のExcelに接続しないことを求めています。
- VBAは、ブックの中のコードを直接書き換えます。文書には、ファイルを正本とした全置換の反映、`.cls`の取り込み、テストの写しが見当たりません。

xlide_mcpは、テストの実行とコンパイルの確認を持つ点で、[設計のテスト](../architecture.md#テスト)に近い候補です。
一方で、ブック（Officeのファイル）を正本とし、書き出した`.bas`と`.cls`を古くなる写しとして扱います。xlwrightの[正本](../requirements.md#正本)の規則とは逆です。

動いているExcelに接続する3つ（ThepExcelMCP、prizmPrograms、Excellm）は、目標の4と合いません。
haris-musaはExcelを使わないため、マクロの実行とテストを扱えません。

## xlwrightで参考になる箇所

- ExcelMcpで読む先は、[ROADMAP](../../ROADMAP.md)の各段階に挙げます。取り入れ方は[ADR 0016](../adr/0016-adopt-excelmcp-designs.md)にあります。
- テストの実行とコンパイルの確認: xlide_mcpと、それが使う`pyVBAharness`。

## 未確認

- ExcelMcpで、動いているExcelへの接続が、どのツールのどの`action`で公開されているか。
- Claude Codeの許可の設定で、1つのツールの中の`action`ごとに許可や拒否を分けられるか。分けられない時は、フックで引数を調べて止める形が候補です。
- xlide_mcpのテストの書き方と、Officeを起動・終了する経路。
