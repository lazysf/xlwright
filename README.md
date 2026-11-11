# xlwright

Windows 版 Excel を PowerShell の COM で操作して、Excel のブックとマクロを作り、確かめるための Claude Code のスキルです。
ファイルを直接書き換えるのではなく、実際の Excel を動かします。

できること:

- シートの値・数式・書式の確認と編集
- VBA（標準・クラス・文書モジュール）と UserForm の書き出し・編集・取り込み
- カスタムリボンの埋め込みと、ボタンを押して動作を確かめること
- VBA プロジェクトをロックした配布用のブックの出力

使い方の正本は [skills/xlwright/SKILL.md](skills/xlwright/SKILL.md) です。

## 前提

- Windows のデスクトップ版 Excel（確かめたのは Microsoft 365 Excel 64-bit）
- Windows PowerShell 5.1（OS 標準）
- Excel のトラストセンターで「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が有効（VBA を扱う時）
- WMI（`Win32_Process.Create`）からプロセスを作れること。Excel を止める見張りのために使います。作れない環境では使えません。
- エージェントの作業中、人はその端末で Excel を使わないこと

## 入れ方

`skills/xlwright/` を、Claude Code の個人スキルの場所 `~/.claude/skills/xlwright/` にコピーします。

```powershell
Copy-Item -Recurse skills\xlwright $HOME\.claude\skills\xlwright
```

更新する時は、古いフォルダを消してからコピーし直します。

## 動作確認

手元の Excel で動くかを確かめます。一時フォルダに新しいブックを作って試し、ほかのファイルは変えません。Excel が動いている時は、何もせずに止まります。

```powershell
powershell.exe -NoProfile -File skills\xlwright\tests\Test-XlWright.ps1
```

`-Full` を付けると、Excel と VBE を画面に出す試験（リボンの押下、配布ブックの出力）も行います。

## 開発

`.claude/skills/xlwright` は、`skills/xlwright` を指す相対パスのシンボリックリンクです。この repo を Claude Code で開くと、`skills/xlwright` がプロジェクトのスキルとして読み込まれ、コピーせずに手直しを試せます。
同じ名前のスキルが `~/.claude/skills/xlwright/` にもあると、個人のスキルが優先され、repo の手直しは使われません。開発中は、個人の場所に入れたコピーを外します。

リンクとして取り出すには、Git の `core.symlinks=true` と、シンボリックリンクを作れる権限（Windows の開発者モードか管理者）が要ります。リンクにならずにテキストファイルとして取り出された時は、それを消して作り直します。

```powershell
cmd /c mklink /D .claude\skills\xlwright ..\..\skills\xlwright
```

開発者モードが有効なら、管理者でなくても `mklink` で作れます。`New-Item -ItemType SymbolicLink` は、開発者モードでも管理者特権を求めて失敗しました（Windows PowerShell 5.1）。

セッションを始めた時に `.claude/skills/` が無かった場合は、`/reload-skills` で読み込み直します。

Markdown の lint には [rumdl](https://rumdl.dev) を使います。Node.js の版は `.nvmrc` にあります。

```powershell
npm ci
npm run lint
```

`npm run lint` は、`lint:md`（rumdl）と `lint:text`（`tools/Test-TextFile.ps1`）を続けて流します。`lint:text` は、Git に入るテキストファイルが BOM なしの UTF-8 で LF だけを使い、制御文字を含まないことと、`.ps1` と `.psm1` が ASCII だけであることを確かめます。`skills/xlwright/tests/data/` の試験データは UTF-8 で読まれるので、ASCII の決まりから外しています。

commit の時に同じ検査を流すには、[prek](https://prek.j178.dev) で Git の hook を入れます。hook の設定は `prek.toml` にあります。

```powershell
prek install
```

### `mcp` ブランチの作業

main（スキル）と `mcp`（MCP サーバー）は、作るものも `.gitignore` も `node_modules` の中身も違います。同じフォルダで切り替えると、前のブランチのビルドの出力や依存が残って混ざります。`mcp` は別のフォルダに worktree として取り出し、そこで作業します。

```powershell
git worktree add ..\xlwright-mcp mcp
```

取り出したフォルダでの準備と確認の手順は、そのフォルダの `docs/development.md` にあります。

## 構成

```text
skills/xlwright/
├── SKILL.md       毎回読む規則と手順
├── references/    作業ごとの詳しい手順（VBA、UserForm、リボン、配布など）
├── scripts/       PowerShell のスクリプト
└── tests/         実際の Excel で確かめる試験と、そのデータ
```

`references/excel-com-rebuild.md` は、別の端末で作った前身のスキル `excel-com` を復元するためのバンドルです。今のスキルは、これを元に作り直しました。中身は今のスキルと一致しないので、使い方は `skills/xlwright/` に従います。

## MCP サーバーについて

このリポジトリでは、もともと Excel を操作する MCP サーバーを開発していました。スキルの形で足りる見込みが立ったため、その開発は中断し、履歴は `mcp` ブランチにあります。再開の手順は、`mcp` ブランチの `TODO.md` にあります。

## ライセンス

[MIT License](LICENSE)
