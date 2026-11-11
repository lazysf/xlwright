---
name: not-excel
description: Excel と関係の無い依頼で、スキルが選ばれないか
tags: [trigger]
max_turns: 10
allowed_tools: [Read, Glob, Grep, Skill]
expected_outcome: xlwright のスキルを呼ばずに、Python のコードで答える。
---

data.csv（列は date と amount）の売上を月ごとに合計する Python のコードを書いて。
