---
name: sheet-check
description: シートの確認と修正の依頼で、スキルが選ばれ、Excel を COM で動かす手順を立てるか
tags: [trigger]
max_turns: 15
allowed_tools: [Read, Glob, Grep, Skill]
expected_outcome: xlwright のスキルを読み、openpyxl などで直接書き換えずに、Excel を COM で開いて確かめ、直した後に開き直して読み戻す手順を示す。
---

C:\work\sales.xlsx の「集計表」シートで、B 列の合計が合っているか確かめて、間違っていたら直してほしい。シェルが使えないなら、どう進めるかの手順だけ教えて。
