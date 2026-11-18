---
name: not-tool
description: 受け取ったブックの集計の確認（ツールの開発でない Excel の依頼）で、スキルが選ばれないか
tags: [trigger]
max_turns: 10
allowed_tools: [Read, Glob, Grep, Skill]
expected_outcome: xlwright のスキルを呼ばずに答える。
---

C:\work\sales.xlsx の「集計表」シートで、B 列の合計が合っているか確かめて、間違っていたら直してほしい。シェルが使えないなら、どう進めるかの手順だけ教えて。
