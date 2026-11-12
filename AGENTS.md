# AGENTS.md

xlwright の開発ルール

## 対象

- このリポジトリは、Claude Code のスキル xlwright（`plugins/xlwright/skills/xlwright/`）を開発し、プラグインとして配布する。使い方は [README.md](README.md)、開発の手順と構成は [docs/development.md](docs/development.md) にある。
- 利用者に写るのは `plugins/xlwright/` の中だけ。開発用のファイルはその外に置く。
- MCP サーバーの開発は中断し、`mcp` ブランチに残してある。`mcp` ブランチの規則は、そのブランチの `AGENTS.md` に従う。

## スキルの書き方

- `SKILL.md` には、毎回の作業で要る規則と手順だけを書く。安全の規則は `SKILL.md` に置く。作業ごとの詳しい手順は `references/` に分け、`SKILL.md` の表から辿れるようにする。
- `${CLAUDE_SKILL_DIR}` が展開されるのは `SKILL.md` の本文だけ。`references/` のコマンド例では、スキルのフォルダを `<skill>` と書く。
- スキルの仕様（frontmatter、変数、読み込まれ方）は、記憶でなく Claude Code の公式 docs で確かめる。
- 文書には、実際の Excel で確かめた事実を書く。確かめていないことは、そうと分かるように書く。

## スクリプト

- Windows PowerShell 5.1 で動かす。PowerShell 7 や `pwsh` を前提にしない。
- `.ps1` と `.psm1` は ASCII だけで書く。日本語のデータは、UTF-8 のデータファイルに分けて読む。
- Excel は `XlWright.psm1` の `Use-ExcelWorkbook`（または `New-ExcelWorkbook`）を通して起動する。`New-Object -ComObject Excel.Application` を直接使わない。
- Excel を止めるのは、その呼び出しが起動した Excel だけにする（PID と起動時刻で照合する）。名前で止めない。
- 実行ポリシーやセキュリティの設定を緩めて続けない。拒否されたら止まって報告する。

## 試験

- スクリプトや手順を変えたら、`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` を流す。リボンや配布の出力に関わる変更では `-Full` も流す。`-Full` は画面に Excel を出すので、先に人へ伝える。
- 新しい振る舞いを加えたら、試験に項目を足す。項目が壊れたものを見つけられるかは、写しのスクリプトをわざと壊して落ちることで確かめる。
- 試験と実験は、一時フォルダの新しいブックか写しで行う。Git で管理するファイルや人のブックを直接変えない。
- Excel が動いている時は、試験を始めずに人へ確かめる。
- ファイルを変えたら、`npm run lint` を流す（Markdown の lint と、文字コード・改行・ASCII の検査）。
- Claude Code を起動して確かめる時（`claude plugin eval`、`claude plugin validate` など）は、Claude Desktop に同梱された CLI を使う。PATH には `claude` が無い。同梱版は `%APPDATA%\Claude-2nd\claude-code\<版>\<ハッシュ>\claude.exe` にあり、アプリの更新で場所が変わるので、使う前に探して、更新日時が最も新しいものを使う。

## 実際の Excel

- 日常の Excel に接続しない。所有を確かめられない Excel を止めない。
- `AccessVBOM` などの Excel やトラストセンターの設定を変えない。必要な時は人に頼む。
- 実験の結果と推測を分ける。未実施・失敗・結果不明を、成功として扱わない。

## 承認

- リポジトリの外のファイルは、読む以外の操作をする前に、対象・変更・戻し方を示して承認を得る。`~/.claude/skills/` へのコピーも含む。セッションの一時ディレクトリは除く。
- commit は人の指示があった時だけ作る。push は人が手元で行う。
- 取り消せない操作（削除、履歴の書き換え）は、対象を示して承認を得てから行う。

## 応答

- 応答と文書は日本語で書く。コードの識別子とコマンドは元のまま書く。
