# TODO

## プラグインとして入れて確かめる

プラグインの検査（`claude plugin validate`）と eval は通っているが、利用者と同じ手順で入れて動かしてはいない。

- 入れ方: この repo を marketplace として足し（`/plugin marketplace add <repo のパス>`）、`/plugin install xlwright@xlwright` で入れる。
- 確かめること: `/xlwright:xlwright` として読み込まれ、スキルの手順（`${CLAUDE_SKILL_DIR}` の展開、`scripts/` の実行）が動くこと。
- 変わるもの: repo の外のユーザー設定（`~/.claude/settings.json` の `extraKnownMarketplaces` と `enabledPlugins`、`~/.claude/plugins/` の marketplace の登録とキャッシュ）。始める前に、対象・変更・戻し方を示して承認を得る。
- 確かめた後: 開発する端末では、プラグイン版を無効にする（`claude plugin disable xlwright@xlwright`）。
- 戻し方: `claude plugin uninstall xlwright@xlwright` と `claude plugin marketplace remove xlwright`。
