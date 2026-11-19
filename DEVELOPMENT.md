# 開発

xlwright を手直しする人のための手順です。利用者向けの入れ方と使い方は [README](README.md) にあります。

## 前提と環境構築

### lint と hook

Markdown の lint には [rumdl](https://rumdl.dev) を使います。Node.js の版は `.nvmrc` にあります。

```powershell
npm ci
```

commit の時に lint と同じ検査を流すには、[prek](https://prek.j178.dev) で Git の hook を入れます。hook の設定は `prek.toml` にあります。

```powershell
prek install
```

`prek install` は、commit の時と push の時の 2 つの hook を入れます。push の時に流れるのは、[版](#リリース)の検査だけです。pre-push の hook を足す前に入れた端末では、`prek install` を流し直します。

### スキルの手直し

`.claude/skills/xlwright` は、`plugins/xlwright/skills/xlwright` を指す相対パスのシンボリックリンクです。この repo を Claude Code で開くと、スキルが `/xlwright` としてプロジェクトから読み込まれ、入れ直さずに手直しを試せます。
プラグイン版も入れている端末では、`/xlwright`（手直し中）と `/xlwright:xlwright`（入れた版）の両方が読み込まれます。開発する端末では、プラグイン版を無効にしておきます。同じ名前のスキルが `~/.claude/skills/xlwright/` にあると、そちらが `/xlwright` として優先されるので、外しておきます。

Codex には、`.agents/skills/xlwright` が同じフォルダを指すリンクとして読まれます。この repo を作業フォルダにして Codex を開くと、スキルが repo のファイルから読み込まれます。プラグインとして入れた版は、写しを読むので、手直しが届くには入れ直しが要ります（[docs/notes/codex-plugin-install.md](docs/notes/codex-plugin-install.md)）。開発する端末では、Codex からもプラグイン版を外しておきます。リンクとプラグイン版を並べた時に、どちらが選ばれるかは確かめていません。

リンクとして取り出すには、Git の `core.symlinks=true` と、シンボリックリンクを作れる権限（Windows の開発者モードか管理者）が要ります。リンクにならずにテキストファイルとして取り出された時は、それを消して作り直します。

```powershell
cmd /c mklink /D .claude\skills\xlwright ..\..\plugins\xlwright\skills\xlwright
```

```powershell
cmd /c mklink /D .agents\skills\xlwright ..\..\plugins\xlwright\skills\xlwright
```

開発者モードが有効なら、管理者でなくても `mklink` で作れます。`New-Item -ItemType SymbolicLink` は、開発者モードでも管理者特権を求めて失敗しました（Windows PowerShell 5.1）。

セッションを始めた時に `.claude/skills/` が無かった場合は、`/reload-skills` で読み込み直します。

## 構成

```text
.claude-plugin/marketplace.json   この repo を Claude Code の marketplace にする定義
.agents/plugins/marketplace.json  この repo を Codex の marketplace にする定義
.agents/skills/xlwright           手直し中のスキルを Codex に読ませるリンク
.claude/skills/xlwright           手直し中のスキルを Claude Code に読ませるリンク
plugins/xlwright/                 配布するプラグイン（この中だけが利用者に写る）
├── plugin.json                   Codex が読む manifest
├── .claude-plugin/plugin.json    Claude Code が読む manifest
├── evals/                        claude plugin eval の case
└── skills/xlwright/
    ├── SKILL.md                  毎回読む規則と手順
    ├── references/               作業ごとの詳しい手順（VBA、UserForm、リボン、配布など）
    ├── scripts/                  PowerShell のスクリプト
    └── tests/                    実際の Excel で確かめる試験と、そのデータ
docs/                             notes（配布しない）
tools/                            開発用の検査（配布しない）
```

`docs/` の中身は [docs/README.md](docs/README.md) から辿れます。

## 変更の確かめ方

### lint

ファイルを変えたら流します。

```powershell
npm run lint
```

`npm run lint` は、`lint:md`（rumdl）と `lint:text`（`tools/Test-TextFile.ps1`）を続けて流します。`lint:text` は、Git に入るテキストファイルが BOM なしの UTF-8 で LF だけを使い、制御文字を含まないことと、`.ps1` と `.psm1` が ASCII だけであることを確かめます。`plugins/xlwright/skills/xlwright/tests/data/` の試験データは UTF-8 で読まれるので、ASCII の決まりから外しています。

### 試験

スクリプトや手順を変えたら、実際の Excel で試験を流します。Excel が動いている時は、何もせずに止まります。

```powershell
powershell.exe -NoProfile -File plugins\xlwright\skills\xlwright\tests\Test-XlWright.ps1
```

- 既定の項目は、Excel を画面に出しません。
- `-Full` を付けると、Excel と VBE を画面に出す項目（コンパイルエラーの検出、マクロが出すダイアログの扱い、リボンの押下と見え方の照合、配布ブックの出力、ブックの見た目の画像）も流します。流している間は、出たウィンドウに触れません。
- `-Case timeout,userform` のように名前を渡すと、その項目だけを流します。`-Full` の項目を名前で選ぶ時も、`-Full` を付けます。
- `-Changed` を付けると、HEAD から変えたスキルのファイル（stage 済み、未 stage、未追跡）に関わる項目だけを流します。どのファイルがどの項目に関わるかは、`Test-XlWright.ps1` の `$CasesByFile` にあります。表に無い `scripts/` のファイル、`tests/data/` の中、`Test-XlWright.ps1` を変えた時は、すべての項目を流します。`-Full` が要る項目は、`-Full` を付けない時は名前を出して飛ばします。`-Case` とは一緒に使えません。

全体を流すと時間がかかります（2026-11-18 に測った時は、`-Full` を付けて 511 秒、既定の項目だけで 307 秒）。多くの項目が Excel を何度も起動するので、項目を選ぶほど短くなります。そこで、次のように流し分けます。

- 作業の途中と commit の前は、`-Changed` で流します。`-Changed` が `-Full` の要る項目を飛ばしても、そのまま commit します。
- push の前に、`-Full` を付けて全体を流します。落ちた項目は直し、直した分を新しい commit にしてから push します。

そのため、push と push の間の commit は、`-Full` の項目で落ちることがあります。`git bisect` などで途中の commit を取り出す時は、そのつもりで扱います。

項目やスクリプトを足した時は、`$CasesByFile` も直します。項目 `changed-selection` が、表の項目の名前と、`-Full` が要る項目の一覧を確かめます。確かめる対象でない準備のブックは、`New-EmptyBook`、`New-BookWithVba`、`New-BookWithMacros` で作ります。どれも最初の 1 回だけ Excel で作り、後は写しを渡します。

### eval

スキルの振る舞いは、`plugins/xlwright/evals/` の case で `claude plugin eval`（Claude Code 2.1.269 以上）にかけて確かめます。各 case は、依頼文（`prompt.md`）と採点の基準（`graders/`）の組です。

| case | 確かめること |
| --- | --- |
| `sheet-check` | ツールのシートの依頼でスキルが選ばれ、Excel を COM で動かす手順を立てる |
| `userform` | UserForm の reference を読み、リボンと配布の reference を読まない |
| `ribbon` | リボンの reference を読み、閉じたブックへ埋め込んで実際に押して確かめる |
| `publish` | 配布の reference を読み、バックアップと `AccessVBOM` の確認を先にし、ロックを読み戻す |
| `vba-import` | VBA の reference を読み、取り込み後にコンパイルを確かめ、ブックにだけあるコンポーネントは名前を示して確認を取ってから消す |
| `safety` | Excel を名前で止めず、`AccessVBOM` を自分で変えない |
| `refused-script` | スクリプトの実行が拒否された時に、Excel を使わない手段で代わりをせず、権限の引き上げを求めるか止まる |
| `not-excel` | Excel と関係の無い依頼でスキルを呼ばない |
| `not-tool` | 受け取ったブックの集計の確認（ツールの開発でない Excel の依頼）でスキルを呼ばない |

```powershell
claude plugin eval plugins/xlwright --case userform --runs 1
claude plugin eval plugins/xlwright
```

- どの case も、許すツールは Read、Glob、Grep、Skill だけです。Excel を実際に動かす確認は `Test-XlWright.ps1` で行います。
- 既定では、1 つの case をプラグインあり・なしで 3 回ずつ流します。会話と同じアカウントで動き、その分の使用量を消費します。最初は `--case` と `--runs 1` で試します。判定の model は `--judge-model sonnet` で指定します。
- 結果は `plugins/xlwright/evals/results/` に書かれ、Git には入れません。
- case は開発用ですが、プラグインの中に置きます。
- シェルを許さない理由、case をプラグインの中に置く理由、原因の分からない FAIL は、[docs/notes/eval.md](docs/notes/eval.md) にあります。
- 上の `claude` には、Claude Desktop に同梱された CLI を使います。同梱版は `%APPDATA%` の下の Claude Desktop のフォルダにある `claude-code\<版>\<ハッシュ>\claude.exe` です。アプリの更新で場所が変わり、古い版も残るので、更新日時が最も新しいものを使います。
- 使用量を Desktop とそろえて把握するため、同梱の CLI は Desktop と同じアカウントでログインして使います。流す前に `claude auth status` を実行し、`loggedIn` が `true`、`authMethod` が `claude.ai`、`email` が Desktop でログインしているアカウントと同じことを確かめます。違う時は、`claude auth login` で Desktop と同じアカウントにログインし直します。

#### Codex

同じ case を、`tools/Invoke-CodexCheck.ps1` で Codex の `codex exec` にかけます。依頼文に `tools/data/codex-check-suffix.txt` の 1 文（ファイルを読むのはよく、Excel を動かすことと書き換えることはしない）を足し、空のフォルダで、読み取りだけの sandbox で流します。Codex に xlwright を入れてから流します（[版を上げる前の確認](#版を上げる前の確認)）。

```powershell
powershell.exe -NoProfile -File tools\Invoke-CodexCheck.ps1
powershell.exe -NoProfile -File tools\Invoke-CodexCheck.ps1 -Case userform
powershell.exe -NoProfile -File tools\Invoke-CodexCheck.ps1 -Case all
```

- 既定で流す case は `userform`、`not-tool`、`refused-script`、`safety` です。1 つの case を 1 回ずつ流します。同梱の `codex.exe` がログインしている ChatGPT のアカウントで動き、その分の使用量を消費します（2026-11-19 に測った時は、スキルを読む case で 1 回あたり入力 8 万〜11 万 token）。
- `codex` には、ChatGPT の app に同梱された `codex.exe`（app のフォルダの `app\resources\codex.exe`）を使います。スクリプトが自分で探します。
- `tool_used` の採点は、記録に残ったコマンドで判定します。`llm` の採点は判定せず、`summary.md` に回答と基準を並べて出します。人か Claude が読んで判定します。
- 記録と `summary.md` は `plugins/xlwright/evals/results/codex/<時刻>/` に書かれ、Git には入れません。`-FromResults <そのフォルダ>` で、Codex を呼ばずに判定だけやり直せます。
- 採点の読み替えと、足す 1 文を今の形にした理由は、[docs/notes/codex-exec.md](docs/notes/codex-exec.md) にあります。

## branch

### `mcp` ブランチ

このリポジトリでは、もともと Excel を操作する MCP サーバーを開発していました。スキルの形で足りる見込みが立ったため、その開発は中断し、履歴は `mcp` ブランチにあります。再開の手順は、`mcp` ブランチの `TODO.md` にあります。

main（スキル）と `mcp`（MCP サーバー）は、作るものも `.gitignore` も `node_modules` の中身も違います。同じフォルダで切り替えると、前のブランチのビルドの出力や依存が残って混ざります。`mcp` は別のフォルダに worktree として取り出し、そこで作業します。

```powershell
git worktree add ..\xlwright-mcp mcp
```

取り出したフォルダでの準備と確認の手順は、そのフォルダの `docs/development.md` にあります。

## リリース

利用者に届く版は、manifest の `version` で決まります。Claude Code は `plugins/xlwright/.claude-plugin/plugin.json` を、Codex は `plugins/xlwright/plugin.json` を読むので、2 つの `version` を同じ値にそろえます。どちらが読まれるかの実測は [docs/notes/codex-plugin-install.md](docs/notes/codex-plugin-install.md) にあります。marketplace の項目（`.claude-plugin/marketplace.json`、`.agents/plugins/marketplace.json`）には `version` を書きません。`version` を変えない限り、何度 push しても利用者は前の版のままです。

利用者に見える変更は、作業中に [CHANGELOG.md](CHANGELOG.md) の `[Unreleased]` へ足していきます。`plugins/xlwright/` を変えて配る時は、push の前に `version` を上げ、同じ commit で `[Unreleased]` の項目を `## [<版>] - <日付>` の節へ移します。作業中の commit ごとに上げる必要はありません。版は semver の形（`0.1.0` など）で書きます。

### 版を上げる前の確認

`version` を上げる commit を作る前に、配る候補を commit した状態で、次を上から順に済ませます。結果は Git に入れません。全部が済んだことを人が確かめてから、版を上げます。`claude plugin eval` と `Invoke-CodexCheck.ps1` は使用量を消費するので、流す前に case の数と回数から起動の回数を見積もります。

1. **静的な検査**: `npm run lint` と、同梱の Claude Code の CLI で `claude plugin validate plugins/xlwright`、`claude plugin validate .` を流します。2 つの manifest の `version` の一致は、push の時の hook が確かめます。
2. **スクリプト**: `Test-XlWright.ps1 -Full` を流します（[試験](#試験)）。エージェントによらないので、1 回でよいです。
3. **Claude Code のスキルの振る舞い**: `claude plugin eval` で、変えた所に関わる case を `--runs 1` で流します（[eval](#eval)）。全部の case は、大きな版の時だけ流します。
4. **Codex に候補を入れる**: 候補の commit を一時フォルダに取り出し、同梱の Codex の CLI で入れます。GitHub から入れた xlwright が Codex に残っていれば、先に外します。marketplace の名前が同じなので、ぶつかります。

   ```powershell
   $codex = Join-Path (Get-AppxPackage -Name OpenAI.Codex).InstallLocation 'app\resources\codex.exe'
   git worktree add --detach "$env:TEMP\xlwright-candidate" HEAD
   & $codex plugin marketplace add "$env:TEMP\xlwright-candidate"
   & $codex plugin add xlwright@xlwright
   ```

5. **Codex のスキルの振る舞い**: `Invoke-CodexCheck.ps1` を流し（[Codex](#codex)）、`summary.md` の `JUDGE` の行を判定します。
6. **Codex で Excel を動かす**: ChatGPT の app の Codex タブで、repo と関係の無い空のフォルダを作業フォルダにしたスレッドを開き、次の依頼を送ります。人が見ている所で行い、スクリプトが sandbox の外での実行を求めること、Excel の計算した値を読み戻して答えることを確かめます。終わった後に Excel のプロセスが残っていないことも見ます。

   ```text
   xlwright スキルを使って、この作業フォルダに book.xlsm を作り、A1 に 1、B1 に =A1+1 を書いて保存し、開き直して B1 の値を答えて。
   ```

7. **Codex から候補を外す**:

   ```powershell
   & $codex plugin remove xlwright@xlwright
   & $codex plugin marketplace remove xlwright
   git worktree remove --force "$env:TEMP\xlwright-candidate"
   ```

8. **push の後**: 版を上げた commit を push した後に、両方で GitHub から入れ直し、新しい版が届くことを見ます。Claude Code は、同梱の CLI の `claude plugin update xlwright@xlwright` の後、Code タブの新しいセッションで `/xlwright:xlwright` が読み込まれることを見ます（手元から入れて push の前に試す手順は、確かめていません）。Codex は、Codex タブの「プラグインマーケットプレイスを追加」に GitHub の `owner/repo` を入れて xlwright を入れ、詳細の画面の版を見ます。

Codex は、入れた時に写しを作ります。候補を直した時は、手順 4 の取り出しと入れ直しからやり直します（[docs/notes/codex-plugin-install.md](docs/notes/codex-plugin-install.md)）。

push の時の hook（`tools/Test-PluginVersion.ps1`）は、origin との差に `plugins/xlwright/` の変更があるのに `version` が origin と同じなら、push を止めます。2 つの manifest の `version` が食い違う時も止めます。origin にまだ無いブランチは、`origin/HEAD`（無ければ `origin/main`）との分岐点と比べます。比べる相手が見つからない時は、警告を出して通します。

版を上げた commit を push した後に、その commit を取り出した状態で `v<版>` の tag を付け、tag を push します。

```powershell
git tag v0.2.0
git push origin v0.2.0
```

tag は、版と commit の対応を残すために付けます。利用者に届く版は変えません。次の版の CHANGELOG を書く時は、`git diff v0.1.0..HEAD -- plugins/xlwright` のように、前の版からの差を見られます。GitHub の Release と CI は使いません。理由は [docs/notes/plugin-install.md](docs/notes/plugin-install.md) にあります。
