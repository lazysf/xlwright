# 開発

xlwright を手直しする人のための手順です。利用者向けの入れ方と使い方は [README](../README.md) にあります。

## スキルの手直し

`.claude/skills/xlwright` は、`plugins/xlwright/skills/xlwright` を指す相対パスのシンボリックリンクです。この repo を Claude Code で開くと、スキルが `/xlwright` としてプロジェクトから読み込まれ、入れ直さずに手直しを試せます。
プラグイン版も入れている端末では、`/xlwright`（手直し中）と `/xlwright:xlwright`（入れた版）の両方が読み込まれます。開発する端末では、プラグイン版を無効にしておきます。同じ名前のスキルが `~/.claude/skills/xlwright/` にあると、そちらが `/xlwright` として優先されるので、外しておきます。

リンクとして取り出すには、Git の `core.symlinks=true` と、シンボリックリンクを作れる権限（Windows の開発者モードか管理者）が要ります。リンクにならずにテキストファイルとして取り出された時は、それを消して作り直します。

```powershell
cmd /c mklink /D .claude\skills\xlwright ..\..\plugins\xlwright\skills\xlwright
```

開発者モードが有効なら、管理者でなくても `mklink` で作れます。`New-Item -ItemType SymbolicLink` は、開発者モードでも管理者特権を求めて失敗しました（Windows PowerShell 5.1）。

セッションを始めた時に `.claude/skills/` が無かった場合は、`/reload-skills` で読み込み直します。

## lint

Markdown の lint には [rumdl](https://rumdl.dev) を使います。Node.js の版は `.nvmrc` にあります。

```powershell
npm ci
npm run lint
```

`npm run lint` は、`lint:md`（rumdl）と `lint:text`（`tools/Test-TextFile.ps1`）を続けて流します。`lint:text` は、Git に入るテキストファイルが BOM なしの UTF-8 で LF だけを使い、制御文字を含まないことと、`.ps1` と `.psm1` が ASCII だけであることを確かめます。`plugins/xlwright/skills/xlwright/tests/data/` の試験データは UTF-8 で読まれるので、ASCII の決まりから外しています。

commit の時に同じ検査を流すには、[prek](https://prek.j178.dev) で Git の hook を入れます。hook の設定は `prek.toml` にあります。

```powershell
prek install
```

## eval

スキルの振る舞いは、`plugins/xlwright/evals/` の case で `claude plugin eval`（Claude Code 2.1.269 以上）にかけて確かめます。各 case は、依頼文（`prompt.md`）と採点の基準（`graders/`）の組です。

| case | 確かめること |
| --- | --- |
| `sheet-check` | シートの依頼でスキルが選ばれ、Excel を COM で動かす手順を立てる |
| `userform` | UserForm の reference を読み、リボンと配布の reference を読まない |
| `ribbon` | リボンの reference を読み、閉じたブックへ埋め込んで実際に押して確かめる |
| `safety` | Excel を名前で止めず、`AccessVBOM` を自分で変えない |
| `not-excel` | Excel と関係の無い依頼でスキルを呼ばない |

```powershell
claude plugin eval plugins/xlwright --case userform --runs 1
claude plugin eval plugins/xlwright
```

- どの case も、許すツールは Read、Glob、Grep、Skill だけです。Windows ではシェルを許した eval が動かないため（シェルは OS の sandbox の中で動き、ネイティブの Windows には sandbox の仕組みが無い）、Excel を実際に動かす確認は `Test-XlWright.ps1` で行います。
- 既定では、1 つの case をプラグインあり・なしで 3 回ずつ流します。会話と同じアカウントで動き、その分の使用量を消費します。最初は `--case` と `--runs 1` で試します。
- 結果は `plugins/xlwright/evals/results/` に書かれ、Git には入れません。

## `mcp` ブランチの作業

main（スキル）と `mcp`（MCP サーバー）は、作るものも `.gitignore` も `node_modules` の中身も違います。同じフォルダで切り替えると、前のブランチのビルドの出力や依存が残って混ざります。`mcp` は別のフォルダに worktree として取り出し、そこで作業します。

```powershell
git worktree add ..\xlwright-mcp mcp
```

取り出したフォルダでの準備と確認の手順は、そのフォルダの `docs/development.md` にあります。

## 構成

```text
.claude-plugin/marketplace.json   この repo を marketplace にする定義
plugins/xlwright/                 配布するプラグイン（この中だけが利用者に写る）
├── .claude-plugin/plugin.json
├── evals/                        claude plugin eval の case
└── skills/xlwright/
    ├── SKILL.md                  毎回読む規則と手順
    ├── references/               作業ごとの詳しい手順（VBA、UserForm、リボン、配布など）
    ├── scripts/                  PowerShell のスクリプト
    └── tests/                    実際の Excel で確かめる試験と、そのデータ
docs/                             開発の手順と notes（配布しない）
tools/                            開発用の検査（配布しない）
```

`plugin.json` には `version` を書かず、利用者は commit を追います。

`references/excel-com-rebuild.md` は、別の端末で作った前身のスキル `excel-com` を復元するためのバンドルです。今のスキルは、これを元に作り直しました。中身は今のスキルと一致しないので、使い方は `plugins/xlwright/skills/xlwright/` に従います。

プラグインとして入れた時の振る舞いを確かめた結果は、[notes/plugin-install.md](notes/plugin-install.md) にあります。

## MCP サーバーについて

このリポジトリでは、もともと Excel を操作する MCP サーバーを開発していました。スキルの形で足りる見込みが立ったため、その開発は中断し、履歴は `mcp` ブランチにあります。再開の手順は、`mcp` ブランチの `TODO.md` にあります。
