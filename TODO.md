# TODO

全体の段階と完了の根拠は [ROADMAP.md](ROADMAP.md) を正本とします。
TODOへの記載は、未承認の設計変更や実機操作を承認するものではありません。

## 直近のTODO

段階0「作り直しの準備」を対象とします。チェックは、記載した完了条件を満たしたときに付けます。
依存は、2→3→5の順です。5は1の結果も使います。1と4は、ほかの項目と独立に進められます。

### 現在の状態

2026-11-02に要件の上流を見直し、作り直しを決めました。

- 旧い設計・計画の文書6つ（`docs/architecture.md`、`docs/api.md`、`docs/spike-0.md`など）は削除しました。消す前の版は`bdd2bb4`にあります。
- 要件の利用の前提と安全性の要件を見直し、新しい[設計](docs/architecture.md)と[API](docs/api.md)を置きました。作り直しの前のADRの扱いは[ADR 0011](docs/adr/0011-retire-pre-rebuild-adrs.md)にあります。
- これらの上流文書は、人の確認を待っています。段階0の項目ではありません。
- `origin`へは一度もpushしていません。mainの全commitはこの端末にしかありません。pushは人がterminalで行います。
- コード（`src/`、`tests/`の`fixtures/`以外、`spikes/Spike0/`、実験用の`scripts/`）は作り直す前のもので、今の`npm run check`が依存しています。項目2で置き換えます。

### 項目

- [ ] **1 AIエージェントのシェルから実Excelを起動できるか確かめる。**
  - 手順:
    1. `spikes/stage0/Test-AgentExcelLaunch.ps1`をPowerShell 5.1で書きます。
    2. スクリプトは、起動の前のEXCELのPIDの一覧を取り、`New-Object -ComObject Excel.Application`で起動します。
    3. `Hwnd`から`GetWindowThreadProcessId`でPIDを引きます。起動の後に増えたEXCELがそのPIDの1つだけの時に限り、自分のものとします。一致しなければ、終了させずに止めます。
    4. `Quit`の後、30秒まで終了を待ちます。残れば、PIDと生成時刻を照合してから終了させます。
    5. 結果（成否、PID、生成時刻、所要時間、例外の種類とHRESULT）を、JSONで標準出力へ出します。
    6. AIエージェントが、sandboxのシェルから`powershell.exe -NoProfile -File`で1回実行します。
  - 守ること: 既に動いているExcelには触れません。この確認は、AGENTS.mdの固定コピーの規則の対象外です。
  - 成立しない時: COMの生成が拒否されたら、迂回せずに記録します。sandboxの外で再試行するかは、人に諮ります。それでも成立しなければ、段階1以降の実Excelのテストは人が起動すると記録します。
  - 完了条件: 結果を`docs/notes/`に記録し、実Excelのテストを誰が起動するかを、[設計](docs/architecture.md#本体の部品の境界)に書きます。
- [ ] **2 旧いコードを外し、空のサーバーと確認の入口を作る。**
  - 外すもの: `src/`、`tests/`（`fixtures/`を除く）、`spikes/Spike0/`、`scripts/`のうち、書き直す`development.ps1`と`inspector.mjs`以外のすべて。`scripts/Show-OwnedExcel.ps1`も外し、段階1でGitの履歴から読み直します。外した後に、[ADR 0009](docs/adr/0009-development-mode-without-excel.md)のStatusを`Superseded`に更新します。新しいサーバーの`--no-excel`が、その判断を置き換えるためです。
  - 作るもの:
    - `src/XlWright.Core/`: `net10.0-windows`のプロジェクト。中身は空です。
    - `src/XlWright.McpServer/`: 引数`--project`と`--no-excel`。ツールは`server_info`だけで、契約は[API](docs/api.md#server_info段階0)のとおりです。ビルドの識別子は、MSBuildの`SourceRevisionId`で埋め込みます。
    - `tests/XlWright.McpServer.Tests/`: ビルドの置き場のサーバーを実プロセスで起動し、`tools/list`と`server_info`を確かめるxUnitのテスト。ツールの定義のスナップショット。
    - `scripts/development.ps1`の書き直し（`npm run check`の入口）: 依存の復元（キャッシュから。不足すれば`-Restore`で公式NuGetから）→ Releaseのビルド → `.xlwright/builds/<build-id>/`へ出す → テスト → Inspector CLIで`tools/list`と`server_info`（`--no-excel`で起動）→ Markdownのチェック → `git diff --check` → 通過を`.xlwright/builds/latest.json`へ記録。古いビルドは、新しい5つと、`latest.json`が指す最後の通過のビルドを残して、フォルダの単位で消します。先にフォルダの名前の変更を試し、失敗したら（動いているサーバーが使っている見込み）飛ばします。
    - `scripts/inspector.mjs`の書き直し: `npm run check`の中では今回のビルドを、`npm run inspect`では`latest.json`のビルドを起動します。
    - `.gitignore`に`plugins/xlwright-dev/.mcp.json`を加えます。
    - `package.json`から、外すスクリプトを指す`prepare:experiment`・`prepare:product`を外します。
    - 結果記録などから外すコードを指すリンクは、`bdd2bb4`の版を指す書き方に直します（[開発ガイド](docs/development.md)の削除済みの文書への参照と同じ書き方）。
  - 依存: 承認済みの版とlock fileを使います（`ModelContextProtocol.Core` 2.3.0など。[環境の確認結果](docs/notes/verification-environment.md#固定した依存)）。新しい依存は加えません。
  - 完了条件: PowerShell 5.1で`npm run check`の全工程が通り、`npm run inspect`で`server_info`の応答を確かめます。
  - `tests/fixtures/`は、段階1で要否を決めるまで残します。
- [ ] **3 開発版のプラグインで、新しい会話だけで新しい版が読まれるか確かめる。**
  - 作るもの: リポジトリの直下の`.claude-plugin/marketplace.json`（`name`は`xlwright-dev`、`owner.name`は`xlwright`、プラグインの`source`は`./plugins/xlwright-dev`）と、`plugins/xlwright-dev/.claude-plugin/plugin.json`（`name`は`xlwright-dev`。`version`は書きません）。`npm run check`が通った時に、`plugins/xlwright-dev/.mcp.json`を書き出す工程を加えます。サーバーの名前は`xlwright`です。中身は、`dotnet`で`${CLAUDE_PLUGIN_ROOT}/../../.xlwright/builds/<build-id>/XlWright.McpServer.dll`を起動し、`--project ${CLAUDE_PROJECT_DIR:-.}`を渡すものです（[パスの制限](docs/architecture.md#パスの制限)）。
  - 人の操作: Claude Codeで`/plugin marketplace add`にこのリポジトリを渡し、`/plugin install xlwright-dev@xlwright-dev`を`local`のスコープ（このリポジトリだけ）で入れます。ユーザーの設定へ書き込むため、人が行います。
  - 確かめる手順:
    1. ビルドAを作ります。人が新しい会話を開いて`server_info`を呼び、Aが返ることを確かめます。
    2. その会話を開いたまま、ビルドBを作ります。ビルドが妨げられないことを確かめます。開いた会話が使うビルドAの置き場の名前の変更が失敗することも確かめます（古いビルドを消す規則の前提）。失敗しなければ、古いビルドを消す工程を止め、消し方を決め直します。
    3. 人が別の新しい会話を開き、Bが返ることを確かめます。1の会話がAを返し続けることも確かめます。
    4. ツールの名前の接頭辞、サーバーが受け取った`--project`の値、作業ディレクトリを記録します。
  - 成立しない時の順: `CLAUDE_CODE_PLUGIN_DIRS`で開発版を読み込む形を試します。次に、版を上げて入れ直す形を試します。それでも読まれなければ、止めて人と相談します。
  - 守ること: リポジトリの外への書き込みは、人が操作するか、操作ごとに承認を得ます。Claude Codeの管理設定のファイルとポリシーのレジストリが無いことは、2026-11-02に標準の場所で確かめました。
  - 完了の後: 入れ方を、[ADR 0003](docs/adr/0003-inspector-centered-development.md)を改めるADRに書きます。調査記録の冒頭の注記を直します。背景は[調査記録](docs/notes/mcp-reload-study.md#プラグインによる開発版の更新2026-11-02)にあります。
- [ ] **4 判定用の題材と確かめる項目の案を作り、人の承認を得る。** [範囲の表](docs/requirements.md#操作)の各行の操作を1回以上使う、練習用の小さな業務アプリです。会話Bの仕様の変更（マクロと数式の両方を直す）の案も含めます。[確かめ方](docs/requirements.md#確かめ方)に従います。
  - 成果物: `docs/acceptance.md`。題材のアプリの仕様、会話Aと会話Bの依頼の文、判定の前に用意する確かめる項目、範囲の表の各行で使う操作の対応、起こし方の手順を書きます。アプリのフォルダは[設計の規約](docs/architecture.md#アプリのフォルダ)に従います。
- [ ] **5 開発ガイドとAGENTS.mdを新しい構成に合わせる。** [開発ガイド](docs/development.md)を書き直し、[AGENTS.md](AGENTS.md)の固定コピー・`npm run check`・Inspector・`project.json`・`XlWright.Core`の規則を直します。固定コピーの規則は、ビルドの置き場から起動する規則に置き換えます。
- [ ] **開発環境の後始末。** 2026-11-01にユーザーの判断で、文書から参照されない生成物を削除しました（`.xlwright/`は約2.1GBから約350MB）。残りは、`.xlwright/spike0/fixture-preparation/`の`s11r2/`と`s14-*/`にある`exported-S11Module.bas`、用途が未確認の`.xlwright/fixture-builder/`と`.xlwright/t11-discarded/`です。これらの扱いはユーザーが判断します。項目2で旧いコードを外す時に、合わせて判断してもらいます。

## 更新方法

作業終了時にTODOを更新し、段階の進捗が変わった場合は [ROADMAP.md](ROADMAP.md) も更新します。
進めない項目には、理由と再開条件を付記します。
段階が完了したら、完了履歴はGitと結果記録に残し、直近のTODOを次の段階へ入れ替えます。
