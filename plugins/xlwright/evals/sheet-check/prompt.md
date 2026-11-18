---
name: sheet-check
description: ツールのシートの編集の依頼で、スキルが選ばれ、Excel を COM で動かす手順を立てるか
tags: [trigger]
max_turns: 15
allowed_tools: [Read, Glob, Grep, Skill]
expected_outcome: xlwright のスキルを読み、openpyxl などで直接書き換えずに、Excel を COM で開いて直し、開き直して読み戻す手順を示す。
---

C:\work\order.xlsm は受注を入力するツールで、「単価」シートに品目と単価の表がある。品目を 1 行足して、「入力」シートの金額の数式がその行も引くように直したい。シェルが使えないなら、どう進めるかの手順だけ教えて。
