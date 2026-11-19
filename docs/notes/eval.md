# この repo で claude plugin eval がどう振る舞うか

- Date: 2026-11-19
- 出典: Claude Desktop に同梱の Claude Code 2.1.326 の `claude plugin eval` の出力と、`plugins/xlwright/evals/results/` の `aggregate-result.json`（2026-11-18 と 2026-11-19 に流した分）。Windows 11。

eval の case を今の形（シェルを許さない、手順を立てる段階を見る）にしている理由と、結果を読む時に知っておくことを書く。流し方は [DEVELOPMENT.md](../../DEVELOPMENT.md) の「eval」を正本とする。

## シェルを許さない理由

Windows では、シェルを許した eval が動かない。シェルは OS の sandbox の中で動き、ネイティブの Windows には sandbox の仕組みが無い。そこで、どの case も許すツールを Read、Glob、Grep、Skill にし、依頼文の最後で「シェルが使えないなら、どう進めるかの手順だけ教えて」と頼む。Excel を実際に動かす確認は、`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` が受け持つ。

## case をプラグインの中に置く理由

`--eval-dir` は、プラグインの外を受け付けなかった。`..` を含むパスは `must stay inside the plugin root`、絶対パスは `must be a relative path inside the plugin` で拒否された。

プラグインを repo の外（scratchpad など）に取り出して流すと、`is not a trusted plugin directory` で止まった。対話のできない実行では、`--trust-plugin` を付けると流れた。取り出したのが自分の repo の commit である時だけ付ける。

## case の選び方

`--case` を 2 回渡す（`--case publish --case sheet-check`）と、後に渡した `sheet-check` だけが流れた（2026-11-19）。エラーも警告も出なかった。`--help` には `--case <glob>` とあるが、複数の case を 1 つの glob で選べるかは試していない。複数の case は、1 つずつ流す。

## 判定の結果の読み方

- `aggregate-result.json` には、grader ごとの票と、判定にかけた最後の回答（`evidence`）が残る。判定の理由は残らない。FAIL の原因は、回答を読んで推し量るしかない。
- 判定の model の既定は haiku で、`publish` では票が割れたことがあった。この repo では `--judge-model sonnet` で流している。
- プラグインあり・なしで比べる時、`tool_used: Skill` のように、プラグインありでしか意味の無い grader は、なしの側で採点されない。

## 原因の分からない FAIL

2026-11-19 の時点で、`ribbon` の `closed-workbook-and-click` は、回答が採点の基準を満たしているように読めるのに FAIL する。プラグインありでも、判定を haiku にしても sonnet にしても、3 票とも FAIL した。原因は確かめていない。

## publish の grader を条件ごとに分けた理由

`publish` は、以前は 1 つの `llm` の grader（`publish-steps`）に 6 つの条件を並べていた。この grader は、回答が基準を満たしているように読めるのに FAIL した。

- sonnet で一度 3 票とも PASS した後、続けて 3 回、3 票とも FAIL した（2026-11-18）。
- 2026-11-19 には、commit `082c510` で haiku と sonnet の判定を 1 回ずつ流し、どちらも FAIL 2 票・PASS 1 票だった。1 つ前の commit `a3e8aab` を sonnet で流すと、3 票とも FAIL だった。
- どの回の回答も、バックアップ、`AccessVBOM` を人に任せること、ウィンドウが出ると先に伝えること、`Publish-Workbook.ps1` での出力、`VBProject.Protection` の読み戻し（配布用が 1、作業ブックが 0）を含み、パスワードを示していなかった。

判定の理由は残らないので、条件が 1 つの grader では、どの条件で落ちたかが分からない。そこで、1 つの条件を 1 つの grader にした（`prepare-first`、`warn-windows`、`use-publish-script`、`read-back-lock`、`no-password`）。rubric は「回答の手順に〜が入っている」と書き、手順だけを示す回答でよいと明記した。`Publish-Workbook.ps1` を使うかは、judge を呼ばない `regex` で見る。

分けた後に sonnet の判定で 1 回流すと（2026-11-19、commit `082c510` の上）、プラグインありは 6 つの grader がすべて通り、`llm` の 4 つは 3 票とも PASS だった。プラグインなし（VBE で人がパスワードを決めてロックする手順）は、6 つすべてが FAIL だった。以前の FAIL が、条件を並べたことと、動作として書いた rubric のどちらによるものかは、確かめていない。

`safety` の `no-kill-by-name` は、1 回 3 票とも FAIL し、流し直すと 3 票とも PASS した。1 回の結果で、規則が抜けたと決めない。

## 既存のマクロを動かさない規則に case が無い理由

既存のマクロを、ユーザーの確認なしに動かさない規則（`SKILL.md` の「マクロが動くことへの注意」）は、単独の case にしていない。試した case は、人から受け取ったブックを確かめる依頼だった。これはスキルの範囲の外で、スキルが選ばれるかが揺れた。プラグインなしでも同じ答えが出て PASS し、規則が抜けたかを見分けられなかった。代わりに、`ribbon` の採点で、既存のマクロを押す前に確認を待つ手順を PASS としている。

## スクリプトのパスに case が無い理由

`SKILL.md` のコマンド例の `<skill>` を、スキルのフォルダの絶対パスに置き換えるかは、eval で見ていない。`sheet-check` に、その絶対パスを示すかを判定する grader を足して 2 回流した（2026-11-18、判定は sonnet）。プラグインありの回答は、どちらも「`<skill>` を `<repo>/plugins/xlwright/skills/xlwright` に置き換える」と正しいフォルダを示していたが、2 回とも 3 票とも FAIL した。2 回目は、置き換え先を示せば PASS と基準に書いた後だった。判定の理由が残らないので原因を確かめられず、grader を外した。パスの解決は、エージェントにシェルを使わせる実機（Codex を含む）で確かめる。

## 未確認

- `ribbon` の FAIL の原因。判定の理由を残す手段は、見つけていない。
- 分けた後の `publish` の grader が、回を重ねても通り続けるか。流したのは 1 回だけ。
- 判定の票のぶれが、どのくらいの頻度で起きるか。
