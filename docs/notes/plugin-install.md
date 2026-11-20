# プラグインとして入れた時の振る舞い

- Date: 2026-11-20
- 出典: Claude Desktop の Code タブの設定の「プラグイン」の画面と `/plugin` の操作、同梱の Claude Code 2.1.321 の `claude plugin list`、`claude plugin marketplace add --help`、`claude plugin update`、`installed_plugins.json` とキャッシュのフォルダ、導入先での `skills/xlwright/tests/Test-XlWright.ps1` の実行（commit `9cf0254`）、Claude Code の docs（<https://code.claude.com/docs/en/plugins>、<https://code.claude.com/docs/en/skills>、<https://code.claude.com/docs/en/discover-plugins>、<https://code.claude.com/docs/en/plugins/loading>、<https://code.claude.com/docs/en/plugins-reference>）、主要なプラグインの repo の manifest と workflow（2026-11-13 の main）、release-please-action の README（<https://github.com/googleapis/release-please-action>）と GitHub の docs（<https://docs.github.com/en/repositories/releasing-projects-on-github/automatically-generated-release-notes>）（2026-11-16）

利用者と同じ手順で入れて確かめた事実と、README の入れ方と版を今の形にしている理由を書く。入れ方そのものは [README](../../README.md) を正本とする。

版として出てくるハッシュ（`689ca7f5a64a`、`cb12836fb035`、`000448a646ae`）は、実際に入った版の文字列をそのまま書いている。commit に署名し直してハッシュが変わったので、今の履歴の commit（順に `9cf0254`、`d3a2067`、`3e730d8`）とは一致しない。中身は同じ。

## 確かめた事実

- Code タブで `/plugin marketplace add owner/repo` を送ると、「ディレクトリ」の画面に「マーケットプレイスを追加」のダイアログが開き、URL の欄に `owner/repo` が入っていた。コマンドとしては実行されない。「同期」を押すと、「Code」の一覧に Xlwright が出た。
- ダイアログの URL の欄は「GitHub の `owner/repo` または Git リポジトリの URL」を求め、手元に clone したフォルダのパスを受け付けなかった。CLI の help には、パスも受け付けると書かれている。
- Xlwright の「+」（インストール）を押すと、user scope で有効になった。版は commit のハッシュ（`689ca7f5a64a`）になった。詳細の画面の「インストール ∨」には、「インストールする」（`~/.claude/`）、「プロジェクト用にインストール（共有）」（`.claude/`）、「プロジェクト用にインストール（個人）」（`.claude.local/`）がある。
- 入れた後に `/plugin install xlwright@xlwright` を送ると、Xlwright の詳細の画面（「削除」と「管理」のボタン）が開いた。
- スキルは `/xlwright:xlwright` として読み込まれた。docs では、プラグインのスキルは常に `/<プラグイン名>:<スキル名>` になる。
- この repo の外のフォルダでは、短い `/xlwright` でもプラグイン版が読み込まれた。Code タブの入力の候補には、`xlwright:xlwright (xlwright)` と出る。この repo の中では、手直し中のスキルも `/xlwright` なので区別できない。
- `${CLAUDE_SKILL_DIR}` は、`~/.claude/plugins/cache/xlwright/xlwright/<版>/skills/xlwright` に展開された。
- 導入先のファイルは、`HEAD` の `plugins/xlwright/` と一覧も中身も一致した。違いは、Claude Code が置く `.in_use/` だけだった。
- 導入先の `skills/xlwright/tests/Test-XlWright.ps1` は、既定の 7 項目も、`-Full` の 10 項目も、すべて合格した。
- Code タブで `/plugin marketplace update xlwright` を送ると、更新せずに設定の「プラグイン」の画面が開いた。
- 設定の「プラグイン」の画面で xlwright の行を押すと詳細が開き、「更新」のボタンと、有効と無効を切り替えるスイッチがある。行の「…」には「無効にする」と「削除」だけがある。スイッチを切ると、`~/.claude/settings.json` の `enabledPlugins` の `xlwright@xlwright` が `false` になった。
- `version` を書いていない時、版は、marketplace の repo の先頭の commit のハッシュ（12 文字）になった。`plugins/xlwright/` を変えていない commit（`d3a2067`、doc だけの変更）でも、キャッシュに版 `cb12836fb035` のフォルダができた（中身は前の版と同じ）。docs の「相対パスのプラグインは、そのディレクトリの commit のハッシュ」（<https://code.claude.com/docs/en/plugins/loading>）とは合わない。
- 新しい commit を push しても、手元の marketplace の写し（`~/.claude/plugins/marketplaces/xlwright`）を取り直すまで、アプリは気づかない。詳細の画面の「更新」は押せず、ツールチップに「最新です」と出る。
- 「ディレクトリ」の画面の「Code」の一覧で marketplace を更新すると、写しが新しい commit になり、新しい版がキャッシュに届いた。しかし入っている版（`~/.claude/plugins/installed_plugins.json`）は切り替わらず、「更新」は押せないままだった。`/reload-plugins` とアプリの再起動でも変わらなかった。新しいセッションも古い版を読み込んだ。
- 同梱の CLI で `claude plugin update xlwright@xlwright` を流すと、入っている版が `689ca7f5a64a` から `000448a646ae` に切り替わった。アプリを再起動すると、新しいセッションは新しい版を読み込んだ。Desktop の画面だけで更新を済ませる手段は、見つからなかった。
- `plugin.json` に `version` を書いた commit（`4373f8a`）を push した後、同梱の CLI で `claude plugin update xlwright@xlwright` を流すと、入っている版が `000448a646ae` から `0.1.0` に切り替わった。キャッシュには版 `0.1.0` のフォルダができ、中身は `plugins/xlwright/` と一致した。前の版のフォルダには `.orphaned_at` が置かれた。続けてもう一度流すと、`already at the latest version (0.1.0)` と出て、何も変わらなかった。
- 2026-11-19 に版を上げて push した後、同梱の CLI で `claude plugin update xlwright@xlwright` を流すと（人の操作）、`installed_plugins.json` の版は push した版、`gitCommitSha` は push した先頭の commit になった。キャッシュのその版のフォルダは、`plugins/xlwright/` と一致した。違いは、Claude Code が置く `.in_use` だけだった（Claude が確かめた）。新しいセッションで `/xlwright:xlwright` が読み込まれたことは、人が確かめた。
- 2026-11-20 に `0.2.0`（commit `0dac807`）を push した後、marketplace を消して入れ直した（人の操作）。`installed_plugins.json` の版は `0.2.0`、`gitCommitSha` は `0dac807`、scope は user になった。キャッシュの `0.2.0` のフォルダは、`plugins/xlwright/` と一致した。違いは、Claude Code が置く `.in_use` と、Git に入らない `evals/results/` だけだった。キャッシュの中で、前に入っていた版のフォルダは消えていた。`0.1.0` とハッシュの名前の 3 つのフォルダは残っていた（どれも Claude が確かめた）。

## 入れ方と版を今の形にしている理由

- スキル名はプラグイン名と同じ `xlwright` にしている。重複は見た目だけで、スキルが選ばれるかは `description` で決まる。名前を変えると、呼び出し名と開発の手順（手直し中の `/xlwright`）が変わる。
- 入れ方は `/plugin marketplace add` と `/plugin install` の 2 段だけを書く。Code タブの画面の操作、更新の手順、自動更新の既定は、Claude Code の使い方の話で xlwright に固有でないので、README に書かない。広く使われているプラグイン（`obra/superpowers`、`anthropics/claude-plugins-official`、`anthropics/skills`、`wshobson/agents`）の README も、`/plugin` のコマンドだけを案内し、scope やチームでの設定に触れていない。
- `claude plugin ...` の CLI のコマンドを入れ方に書かない。Desktop アプリは CLI を同梱するが、PATH に `claude` が無い。
- scope（user、project、local）で入れ方を分けない。`claude plugin install` と `claude plugin marketplace add` は `--scope` を持ち、local にすれば `~/.claude/settings.json` に書かずに repo ごとに入れられる。ただし、どの scope でも `~/.claude/plugins/` には書き込む。
- `plugin.json` に `version` を書き、上げた時だけ新しい版が利用者に届くようにする。書かないと、`plugins/xlwright/` を変えていない commit でも新しい版ができる（上の事実）。docs（plugins/loading）は、manifest で版を固定すると、その文字列を変えるまで何度 push しても利用者はキャッシュの写しのままと書いている。
- marketplace の項目には `version` を書かない。docs（plugins-reference）では manifest の `version` が項目の `version` より優先されるので、項目に書いても使われず、食い違いの元になるだけになる。
- 版の上げ忘れは、pre-push の hook（`tools/Test-PluginVersion.ps1`）で止める。commit ごとには確かめず、作業中は何度 commit してもよく、配る直前に 1 回だけ版を上げる。`wshobson/agents` も、変えたプラグインの版が上がっているかを確かめるが、場所は PR の CI になっている。この repo は main へ直接 push するので、push の時に確かめる。
- 版を CI で自動で付ける形（release-please など）は採らない。調べた主要な repo（`obra/superpowers`、`wshobson/agents`、`anthropics/claude-plugins-official`、`anthropics/claude-code`、`anthropics/skills`）は、どれも版を手で書くか、書かずにハッシュに任せていた。自動にすると、新しい依存と、Actions に PR を作らせる権限の設定が要る。docs の版の決め方（manifest の `version`、項目の `version`、commit のハッシュの順）に git のタグは出てこない。
- release-please を採らない理由は、ほかに 2 つある。release-please は版を上げる PR を自分で作るので、版を上げていない push を止める pre-push の hook と両立しない。CHANGELOG は commit message から作られるので、利用者に向けて日本語でまとめて書く今の形にならない。
- tag は、版を上げた push の後に手で付ける。版と commit の対応を残すためで、利用者の版には関わらない。GitHub の Release と CI は足さない。Release の中身は CHANGELOG.md の写しになる。lint は commit の時の hook で流れているので、CI で流し直さない。
- 公式のプラグインの一覧（plugin directory）には申請しない。人と決めた判断（2026-11-19）。そのため、`claude plugin validate plugins/xlwright` が出す一覧向けの確認は、直さずに残している（同梱の Claude Code 2.1.326、2026-11-19）。
  - 警告 `ICON_MISSING`: プラグインに icon が無い。直すには、`.claude-plugin/icon.png` に正方形の PNG を置くか、`plugin.json` の `icon` に正方形の PNG か JPEG を書く。一辺が 512〜2048 px で、2 MB 未満にする。SVG と WebP は受け付けない。一覧の icon になるのは、developer portal で初めて保存または申請した時のファイルだけで、後から変えても一覧の icon は変わらない。
  - 注記: プラグインのフォルダに README が無いので、repo の一番上の README.md が使われる。申請は commit したものだけを、repo の一番上から読む。
  - どちらも validate の結果（`Validation passed`）を変えない。

## 未確認

- CLI の `claude plugin marketplace add` に、手元に clone したフォルダのパスを渡す形。
- 入れる前に、Code タブで `/plugin install xlwright@xlwright` を送る形。
- `marketplace add`、`install`、`marketplace update` のほかの `/plugin` のサブコマンドを、Code タブで送る形。
- 詳細の画面の「更新」が押せるようになる条件。marketplace を取り直した後も、再起動した後も押せなかった。`version` を上げた時に押せるようになるかは見ていない（`0.1.0` への切り替えは CLI で行った）。
- 入っている版より小さい版を push した時に、`claude plugin update` で切り替わるか。2026-11-20 は、marketplace を消して入れ直した。
- Code タブで、marketplace の自動更新を有効にする操作があるか。docs は `/plugin` の Marketplaces の画面の「Enable auto-update」を挙げるが、Desktop の画面では見ていない。
- CI でタグを付けるだけで、利用者の版が変わらないか。docs の版の決め方にタグが出てこないことからの推測で、試していない。
- GitHub が自動で作る Release のノートに、PR を通さずに main へ push した commit が載るか。docs は merge した PR だけを挙げている。載らなければ、main へ直接 push するこの repo では使えない。
