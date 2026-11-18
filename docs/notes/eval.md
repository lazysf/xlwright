# この repo で claude plugin eval がどう振る舞うか

- Date: 2026-11-18
- 出典: Claude Desktop に同梱の Claude Code 2.1.326 の `claude plugin eval` の出力と、`plugins/xlwright/evals/results/` の `aggregate-result.json`（2026-11-18 に流した分）。Windows 11。

eval の case を今の形（シェルを許さない、手順を立てる段階を見る）にしている理由と、結果を読む時に知っておくことを書く。流し方は [DEVELOPMENT.md](../../DEVELOPMENT.md) の「eval」を正本とする。

## シェルを許さない理由

Windows では、シェルを許した eval が動かない。シェルは OS の sandbox の中で動き、ネイティブの Windows には sandbox の仕組みが無い。そこで、どの case も許すツールを Read、Glob、Grep、Skill にし、依頼文の最後で「シェルが使えないなら、どう進めるかの手順だけ教えて」と頼む。Excel を実際に動かす確認は、`plugins/xlwright/skills/xlwright/tests/Test-XlWright.ps1` が受け持つ。

## case をプラグインの中に置く理由

`--eval-dir` は、プラグインの外を受け付けなかった。`..` を含むパスは `must stay inside the plugin root`、絶対パスは `must be a relative path inside the plugin` で拒否された。

プラグインを repo の外（scratchpad など）に取り出して流すと、`is not a trusted plugin directory` で止まった。対話のできない実行では、`--trust-plugin` を付けると流れた。取り出したのが自分の repo の commit である時だけ付ける。

## 判定の結果の読み方

- `aggregate-result.json` には、grader ごとの票と、判定にかけた最後の回答（`evidence`）が残る。判定の理由は残らない。FAIL の原因は、回答を読んで推し量るしかない。
- 判定の model の既定は haiku で、`publish` では票が割れたことがあった。この repo では `--judge-model sonnet` で流している。
- プラグインあり・なしで比べる時、`tool_used: Skill` のように、プラグインありでしか意味の無い grader は、なしの側で採点されない。

## 原因の分からない FAIL

2026-11-18 の時点で、次の 2 つは、回答が採点の基準を満たしているように読めるのに FAIL する。原因は確かめていない。

- `ribbon` の `closed-workbook-and-click`: プラグインありでも、判定を haiku にしても sonnet にしても、3 票とも FAIL した。
- `publish` の `publish-steps`: sonnet で一度 3 票とも PASS した後、続けて 3 回、3 票とも FAIL した。うち 1 回は、`description` をツールの開発に絞る前の写しで流したもので、絞ったことによる FAIL ではない。

`safety` の `no-kill-by-name` は、1 回 3 票とも FAIL し、流し直すと 3 票とも PASS した。1 回の結果で、規則が抜けたと決めない。

## 既存のマクロを動かさない規則に case が無い理由

既存のマクロを、ユーザーの確認なしに動かさない規則（`SKILL.md` の「マクロが動くことへの注意」）は、単独の case にしていない。試した case は、人から受け取ったブックを確かめる依頼だった。これはスキルの範囲の外で、スキルが選ばれるかが揺れた。プラグインなしでも同じ答えが出て PASS し、規則が抜けたかを見分けられなかった。代わりに、`ribbon` の採点で、既存のマクロを押す前に確認を待つ手順を PASS としている。

## 未確認

- `ribbon` と `publish` の FAIL の原因。判定の理由を残す手段は、見つけていない。
- 判定の票のぶれが、どのくらいの頻度で起きるか。
