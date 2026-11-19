# Codex で手元のフォルダからプラグインを入れた時の振る舞い

- Date: 2026-11-19
- 出典: Codex の desktop app（`OpenAI.Codex` 26.1111.2228.0）の「プラグインマーケットプレイスを追加」の画面とプラグインの詳細の画面の操作（人が行った）、`~/.codex/config.toml` の `[marketplaces.*]` と `~/.codex/plugins/cache/` の読み取り（Claude が行った）、公式 docs の [Package your plugin](https://developers.openai.com/plugins/build/plugins)。Windows 11。試したのは commit `386725f`。

版を上げて push する前に、Codex で候補を試すための事実を書く。GitHub から入れた時の振る舞いと、sandbox の中での振る舞いは [codex-sandbox.md](codex-sandbox.md) にある。

## 確かめた事実

- 追加の画面は、ソースに GitHub の repo、Git の URL、手元のフォルダを受け付ける。ソースに repo のルート（`.claude-plugin/marketplace.json` がある階層）の絶対パスを入れ、Git ref とスパースパスを空のままにすると、marketplace `xlwright` として足された（人の操作）。
- 足した後の `config.toml` には、`[marketplaces.xlwright]` が `source_type = "local"` で 1 件だけあった。前に GitHub から足した `xlwright` の項目は残っていなかった。足した時に置き換わったのか、その前に外されていたのかは確かめていない。
- プラグインの詳細の画面には、バージョン（`plugin.json` の `version`）、マーケットプレイス `xlwright`、ソースとして `plugins/xlwright` の絶対パスが出た（人の画面の写し）。
- 写しは `~/.codex/plugins/cache/xlwright/xlwright/<版>/` に置かれた。docs は手元のプラグインの版を `local` と書くが、実機では `plugin.json` の `version` のフォルダだった。中身は `plugins/xlwright/` と一致した（`diff -rq`）。
- 一時フォルダに detached の worktree を作り、`plugins/xlwright/` に目印のファイル（中身 `marker-1`）を足した。marketplace をその worktree に差し替えて入れ直すと、同じ版のフォルダが更新され、目印が届いた。版が同じでも、入れ直せば写しは新しくなる。
- 続けて目印を `marker-2` に書き換え、入れ直さずに app を再起動した。写しは `marker-1` のままで、フォルダの更新日時も変わらなかった。docs の「再起動で新しいファイルを読み込む」とは合わない。

## 試す時の手順に効くこと

- 手直しを Codex に届けるには、入れ直す。再起動だけでは届かない。
- 写しのフォルダの名前は版で決まり、手元の候補と GitHub の版を区別しない。Codex で試した候補は、commit のハッシュで記録する。
- 手元のフォルダを足すと、commit していない変更も写しに入る。候補の commit だけを試すには、worktree に取り出して足す。

## プラグインを入れずに手直し中のスキルを読ませる

docs（[Build skills](https://learn.chatgpt.com/docs/build-skills.md)）は、起動したフォルダから repo のルートまでの各階層の `.agents/skills` を探し、リンクしたスキルのフォルダはリンク先をたどる、と書いている。

repo のルートに `.agents/skills/xlwright`（`../../plugins/xlwright/skills/xlwright` を指すシンボリックリンク、`mklink /D` で作った）を置いて確かめた（2026-11-19）。この時、Codex からは xlwright の marketplace もプラグインも外してあった（人の報告。`config.toml` に xlwright の marketplace が無いことは Claude が確かめた）。

- 作業フォルダを repo にした新しいスレッドで `/xl` と打つと、候補は「Xlwright: Xlwright」の 1 件だけだった。右端に `xlwright` と出た（人の画面の写し）。
- 「xlwright スキルの SKILL.md の絶対パスを答えて」と頼むと、エージェントは `.agents/skills/xlwright/SKILL.md` の絶対パスを答えた（人が受け取った回答）。

答えたパスがリンクの下だったので、リンクのスキルは repo のファイルを読んでいると見ている。手直しのたびに入れ直す必要は無い見込みだが、手直しが届くかは試していない。

## Claude 用と Codex 用の manifest を並べた時

docs（[Package your plugin](https://developers.openai.com/plugins/build/plugins)）は、どのエージェントでも読める形としてプラグインのフォルダの直下の `plugin.json` を勧め、marketplace は `.agents/plugins/marketplace.json`（`.claude-plugin/marketplace.json` は legacy-compatible）に置くと書く。両方を置いた時にどちらを使うかは、書いていない。

一時フォルダの worktree（commit `f6dfc01`）に、`plugins/xlwright/plugin.json` と `.agents/plugins/marketplace.json` を足して確かめた（2026-11-19）。見分けるため、それぞれの `interface.displayName` を `xlwright (manifest test)` と `xlwright (agents marketplace)` にした。足す前に、Codex から xlwright の marketplace とプラグインは外してあった（人の報告）。

- app で worktree を marketplace として足すと、「マーケットプレイス」の一覧には `xlwright (agents marketplace)` の 1 件だけが「1個のプラグイン」として出た（人の画面の写し）。同じフォルダに `.claude-plugin/marketplace.json` もあるが、`.agents/plugins/marketplace.json` が使われ、marketplace は 2 件にならなかった。
- プラグインの一覧と詳細の画面には、`xlwright (manifest test)`、開発者 `lazysf`、カテゴリ `Productivity`、バージョン（`plugin.json` の `version`）が出た（人の画面の写し）。表示名、開発者、カテゴリは、直下の `plugin.json` の `extensions.com.openai.interface` の値だった。
- `config.toml` には `[marketplaces.xlwright]`（worktree を指す）と `[plugins."xlwright@xlwright"]` が 1 件ずつあった。写しの `cache/xlwright/xlwright/<版>/` には、直下の `plugin.json` と `.claude-plugin/` の両方が入っていた（Claude が読み取りで確かめた）。
- 同じ worktree で、同梱の Claude Code 2.1.326 の `claude plugin validate` を、プラグインのフォルダと repo のルートに流すと、どちらも通った。読まれたのは `.claude-plugin/plugin.json` と `.claude-plugin/marketplace.json` で、出た警告（アイコンが無い）は足す前の repo と同じだった。

## 未確認

- 候補の右端の `xlwright` が何を表すか。プラグインを外していたので、project の名前と見ているが確かめていない。
- プラグイン版の xlwright も入れた時に、リンクのスキルと 2 件並ぶか。docs は、同じ `name` のスキルを合わせずに両方出す、と書く。
- `SKILL.md` の `description` を変えた時に、再起動せずに候補へ反映されるか。
- 入れ直しの代わりに、marketplace の更新（`codex plugin marketplace upgrade` や画面の操作）で写しが新しくなるか。
- `codex exec` が、app で入れたプラグインのスキルを読み込むか。
- GitHub の marketplace と手元の marketplace を、名前を分けずに並べて置けるか。今回は、足した後に GitHub の項目が残っていなかった。
