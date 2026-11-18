# スキルの description を今の形にしている理由

- Date: 2026-11-18
- 出典: Claude Code の docs（<https://code.claude.com/docs/en/skills>）と Agent Skills の docs（<https://platform.claude.com/docs/en/agents-and-tools/agent-skills/best-practices>）、`not-tool` の case を同梱の Claude Code 2.1.326 で流した結果（commit `98358d0`）

`description` の本文は `plugins/xlwright/skills/xlwright/SKILL.md` を正本とする。ここには、その形にしている理由を書く。

## 目的と、使わない場面を書く理由

xlwright の目的は、Excel を土台にしたツール（マクロ、UserForm、カスタムリボンを持ち、人が操作して使うブック）の開発である（AGENTS.md）。通常のブックの分析や整形は、`xlsx` スキルなど、ほかの表計算のスキルの役目になる。

`description` が「シートの値・数式・書式」を先に挙げていると、受け取ったブックの集計を確かめる依頼でもスキルが選ばれた。そこで、`description` に目的を先に書き、使わない場面（受け取ったブックの分析、集計の確認、データの整形や変換）も書いている。

`not-tool`（`C:\work\sales.xlsx` の集計が合っているか確かめる依頼）を、プラグインありで 3 回ずつ流した結果（2026-11-18、model は `claude-opus-5-5`）:

| `description` | スキルが選ばれた回数 |
| --- | --- |
| 「シートの値・数式・書式」を先に挙げた形（`98358d0` の前） | 3 回のうち 3 回 |
| 今の形 | 3 回のうち 0 回 |

今の形でも、ツールのシートの編集の依頼（`sheet-check`）では、スキルが選ばれた（1 回）。スキルが選ばれたかを数える grader は、`sheet-check` にしか無い。

eval を流す環境に、`xlsx` などほかのスキルが入っているかは確かめていない。入っていなければ、Excel の依頼で選べるスキルは xlwright だけになる。その環境でも選ばれないので、`description` の書き方だけで選ばれる範囲を絞れている。

## 書かないこと

- 範囲の外の依頼を断る手順は、`SKILL.md` に書かない。範囲の外では、スキルがそもそも選ばれないようにする。
- ほかの表計算のスキルの側は直さない。この repo の外のものなので。

## 書く時の制約

- Claude Code は、スキルの一覧で `description` と `when_to_use` を合わせて 1,536 文字で切る。主な用途を先に書く。
- `<` と `>` は、`&lt;` と `&gt;` と表示される。
- Agent Skills の決まりでは、`description` は 1,024 文字まで、XML のタグを含めない。何をするかと、いつ使うかを書く。
