# プラグインとして入れた時の振る舞い

- Date: 2026-11-12
- 出典: Claude Desktop の Code タブの設定の「プラグイン」の画面と `/plugin` の操作、同梱の Claude Code 2.1.321 の `claude plugin list` と `claude plugin marketplace add --help`、導入先での `tests/Test-XlWright.ps1` の実行（commit `689ca7f`）、Claude Code の docs（<https://code.claude.com/docs/en/plugins>、<https://code.claude.com/docs/en/skills>、<https://code.claude.com/docs/en/discover-plugins>）

利用者と同じ手順で入れて確かめた事実と、README の入れ方を今の形にしている理由を書く。入れ方そのものは [README](../../README.md) を正本とする。

## 確かめた事実

- Code タブで `/plugin marketplace add owner/repo` を送ると、「ディレクトリ」の画面に「マーケットプレイスを追加」のダイアログが開き、URL の欄に `owner/repo` が入っていた。コマンドとしては実行されない。「同期」を押すと、「Code」の一覧に Xlwright が出た。
- ダイアログの URL の欄は「GitHub の `owner/repo` または Git リポジトリの URL」を求め、手元に clone したフォルダのパスを受け付けなかった。CLI の help には、パスも受け付けると書かれている。CLI でパスを渡す形は試していない。
- Xlwright の「+」（インストール）を押すと、user scope で有効になった。詳細の画面の「インストール ∨」には、「インストールする」（`~/.claude/`）、「プロジェクト用にインストール（共有）」（`.claude/`）、「プロジェクト用にインストール（個人）」（`.claude.local/`）がある。
- 入れた後に `/plugin install xlwright@xlwright` を送ると、Xlwright の詳細の画面（「削除」と「管理」のボタン）が開いた。入れる前に送る形は試していない。版は commit のハッシュ（`689ca7f5a64a`）になった。
- スキルは `/xlwright:xlwright` として読み込まれた。docs では、プラグインのスキルは常に `/<プラグイン名>:<スキル名>` になる。
- この repo の外のフォルダでは、短い `/xlwright` でもプラグイン版が読み込まれた。Code タブの入力の候補には、`xlwright:xlwright (xlwright)` と出る。この repo の中では、手直し中のスキルも `/xlwright` なので区別できない。
- `${CLAUDE_SKILL_DIR}` は、`~/.claude/plugins/cache/xlwright/xlwright/<版>/skills/xlwright` に展開された。
- 導入先のファイルは、`HEAD` の `plugins/xlwright/` と一覧も中身も一致した。違いは、Claude Code が置く `.in_use/` だけだった。
- 導入先の `tests/Test-XlWright.ps1` は、既定の 7 項目も、`-Full` の 10 項目も、すべて合格した。
- Code タブで `/plugin marketplace update xlwright` を送ると、更新せずに設定の「プラグイン」の画面が開いた。ほかの `/plugin` のサブコマンドを Code タブで送る形は試していない。
- 設定の「プラグイン」の画面で xlwright の行を押すと詳細が開き、「更新」のボタンと、有効と無効を切り替えるスイッチがある。行の「…」には「無効にする」と「削除」だけがある。入れた版から GitHub に新しい commit が無い時、「更新」は押せなかった。

## README の入れ方を今の形にしている理由

- スキル名はプラグイン名と同じ `xlwright` のままにした。重複は見た目だけで、スキルが選ばれるかは `description` で決まる。名前を変えると、呼び出し名と開発の手順（手直し中の `/xlwright`）が変わる。
- 入れ方は `/plugin marketplace add` と `/plugin install` の 2 段を書き、Code タブではその 2 つが画面を開くだけなので、画面での操作を添える。広く使われているプラグイン（`obra/superpowers`、`anthropics/claude-plugins-official`、`anthropics/skills`、`wshobson/agents`）の README も、`/plugin` のコマンドだけを案内し、scope やチームでの設定に触れていない。
- `claude plugin ...` の CLI のコマンドを入れ方に書かない。Desktop アプリは CLI を同梱するが、PATH に `claude` が無い。
- scope（user、project、local）で入れ方を分けない。`claude plugin install` と `claude plugin marketplace add` は `--scope` を持ち、local にすれば `~/.claude/settings.json` に書かずに repo ごとに入れられる。ただし、どの scope でも `~/.claude/plugins/` には書き込む。

## 未確認

- GitHub に新しい commit を push した後、詳細の画面の「更新」が押せるようになり、押すと新しい版に変わるか。README の更新の手順は、CLI のコマンド（`claude plugin marketplace update`、`claude plugin update`）で書いている。
