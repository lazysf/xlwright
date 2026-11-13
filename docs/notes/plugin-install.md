# プラグインとして入れた時の振る舞い

- Date: 2026-11-13
- 出典: Claude Desktop の Code タブの設定の「プラグイン」の画面と `/plugin` の操作、同梱の Claude Code 2.1.321 の `claude plugin list`、`claude plugin marketplace add --help`、`claude plugin update`、`installed_plugins.json` とキャッシュのフォルダ、導入先での `tests/Test-XlWright.ps1` の実行（commit `689ca7f`）、Claude Code の docs（<https://code.claude.com/docs/en/plugins>、<https://code.claude.com/docs/en/skills>、<https://code.claude.com/docs/en/discover-plugins>、<https://code.claude.com/docs/en/plugins/loading>、<https://code.claude.com/docs/en/plugins-reference>）、主要なプラグインの repo の manifest と workflow（2026-11-13 の main）

利用者と同じ手順で入れて確かめた事実と、README の入れ方と版を今の形にしている理由を書く。入れ方そのものは [README](../../README.md) を正本とする。

## 確かめた事実

- Code タブで `/plugin marketplace add owner/repo` を送ると、「ディレクトリ」の画面に「マーケットプレイスを追加」のダイアログが開き、URL の欄に `owner/repo` が入っていた。コマンドとしては実行されない。「同期」を押すと、「Code」の一覧に Xlwright が出た。
- ダイアログの URL の欄は「GitHub の `owner/repo` または Git リポジトリの URL」を求め、手元に clone したフォルダのパスを受け付けなかった。CLI の help には、パスも受け付けると書かれている。CLI でパスを渡す形は試していない。
- Xlwright の「+」（インストール）を押すと、user scope で有効になった。版は commit のハッシュ（`689ca7f5a64a`）になった。詳細の画面の「インストール ∨」には、「インストールする」（`~/.claude/`）、「プロジェクト用にインストール（共有）」（`.claude/`）、「プロジェクト用にインストール（個人）」（`.claude.local/`）がある。
- 入れた後に `/plugin install xlwright@xlwright` を送ると、Xlwright の詳細の画面（「削除」と「管理」のボタン）が開いた。入れる前に送る形は試していない。
- スキルは `/xlwright:xlwright` として読み込まれた。docs では、プラグインのスキルは常に `/<プラグイン名>:<スキル名>` になる。
- この repo の外のフォルダでは、短い `/xlwright` でもプラグイン版が読み込まれた。Code タブの入力の候補には、`xlwright:xlwright (xlwright)` と出る。この repo の中では、手直し中のスキルも `/xlwright` なので区別できない。
- `${CLAUDE_SKILL_DIR}` は、`~/.claude/plugins/cache/xlwright/xlwright/<版>/skills/xlwright` に展開された。
- 導入先のファイルは、`HEAD` の `plugins/xlwright/` と一覧も中身も一致した。違いは、Claude Code が置く `.in_use/` だけだった。
- 導入先の `tests/Test-XlWright.ps1` は、既定の 7 項目も、`-Full` の 10 項目も、すべて合格した。
- Code タブで `/plugin marketplace update xlwright` を送ると、更新せずに設定の「プラグイン」の画面が開いた。ほかの `/plugin` のサブコマンドを Code タブで送る形は試していない。
- 設定の「プラグイン」の画面で xlwright の行を押すと詳細が開き、「更新」のボタンと、有効と無効を切り替えるスイッチがある。行の「…」には「無効にする」と「削除」だけがある。スイッチを切ると、`~/.claude/settings.json` の `enabledPlugins` の `xlwright@xlwright` が `false` になった。
- `version` を書いていない時、版は、marketplace の repo の先頭の commit のハッシュ（12 文字）になった。`plugins/xlwright/` を変えていない commit（`cb12836`、README だけの変更）でも、キャッシュに版 `cb12836fb035` のフォルダができた（中身は前の版と同じ）。docs の「相対パスのプラグインは、そのディレクトリの commit のハッシュ」（<https://code.claude.com/docs/en/plugins/loading>）とは合わない。
- 新しい commit を push しても、手元の marketplace の写し（`~/.claude/plugins/marketplaces/xlwright`）を取り直すまで、アプリは気づかない。詳細の画面の「更新」は押せず、ツールチップに「最新です」と出る。
- 「ディレクトリ」の画面の「Code」の一覧で marketplace を更新すると、写しが新しい commit になり、新しい版がキャッシュに届いた。しかし入っている版（`~/.claude/plugins/installed_plugins.json`）は切り替わらず、「更新」は押せないままだった。`/reload-plugins` とアプリの再起動でも変わらなかった。新しいセッションも古い版を読み込んだ。
- 同梱の CLI で `claude plugin update xlwright@xlwright` を流すと、入っている版が `689ca7f5a64a` から `000448a646ae` に切り替わった。アプリを再起動すると、新しいセッションは新しい版を読み込んだ。Desktop の画面だけで更新を済ませる手段は、見つからなかった。

## 入れ方と版を今の形にしている理由

- スキル名はプラグイン名と同じ `xlwright` のままにした。重複は見た目だけで、スキルが選ばれるかは `description` で決まる。名前を変えると、呼び出し名と開発の手順（手直し中の `/xlwright`）が変わる。
- 入れ方は `/plugin marketplace add` と `/plugin install` の 2 段だけを書く。Code タブの画面の操作、更新の手順、自動更新の既定は、Claude Code の使い方の話で xlwright に固有でないので、README に書かない。広く使われているプラグイン（`obra/superpowers`、`anthropics/claude-plugins-official`、`anthropics/skills`、`wshobson/agents`）の README も、`/plugin` のコマンドだけを案内し、scope やチームでの設定に触れていない。
- `claude plugin ...` の CLI のコマンドを入れ方に書かない。Desktop アプリは CLI を同梱するが、PATH に `claude` が無い。
- scope（user、project、local）で入れ方を分けない。`claude plugin install` と `claude plugin marketplace add` は `--scope` を持ち、local にすれば `~/.claude/settings.json` に書かずに repo ごとに入れられる。ただし、どの scope でも `~/.claude/plugins/` には書き込む。
- `plugin.json` に `version` を書き、上げた時だけ新しい版が利用者に届くようにする。書かないと、`plugins/xlwright/` を変えていない commit でも新しい版ができる（上の事実）。docs（plugins/loading）は、manifest で版を固定すると、その文字列を変えるまで何度 push しても利用者はキャッシュの写しのままと書いている。
- marketplace の項目には `version` を書かない。docs（plugins-reference）では manifest の `version` が項目の `version` より優先されるので、項目に書いても使われず、食い違いの元になるだけになる。
- 版の上げ忘れは、pre-push の hook（`tools/Test-PluginVersion.ps1`）で止める。commit ごとには確かめず、作業中は何度 commit してもよく、配る直前に 1 回だけ版を上げる。`wshobson/agents` も、変えたプラグインの版が上がっているかを確かめるが、場所は PR の CI になっている。この repo は main へ直接 push するので、push の時に確かめる。
- 版を CI で自動で付ける形（release-please など）は採らない。調べた主要な repo（`obra/superpowers`、`wshobson/agents`、`anthropics/claude-plugins-official`、`anthropics/claude-code`、`anthropics/skills`）は、どれも版を手で書くか、書かずにハッシュに任せていた。自動にすると、新しい依存と、Actions に PR を作らせる権限の設定が要る。docs の版の決め方（manifest の `version`、項目の `version`、commit のハッシュの順）に git のタグは出てこないので、CI でタグを付けるだけでは利用者の版は変わらないと見ている。

## 未確認

- 詳細の画面の「更新」が押せるようになる条件。marketplace を取り直した後も、再起動した後も押せなかった。
- `version` を書いた版（`0.1.0`）を利用者が受け取る時の振る舞い。ハッシュの版から `0.1.0` へ `claude plugin update` で切り替わるか、版を上げた後に詳細の画面の「更新」が押せるようになるかは、push の後に確かめる。
- Code タブで、marketplace の自動更新を有効にする操作があるか。docs は `/plugin` の Marketplaces の画面の「Enable auto-update」を挙げるが、Desktop の画面では見ていない。
