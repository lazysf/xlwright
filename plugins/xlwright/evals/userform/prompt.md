---
name: userform
description: UserForm の依頼で、UserForm の reference を読み、関係の無い reference を読まないか
tags: [reference]
max_turns: 15
allowed_tools: [Read, Glob, Grep, Skill]
expected_outcome: xlwright のスキルを読んだ後、references/userform.md を読む。リボンや配布の reference は読まない。
---

C:\work\input.xlsm に、顧客名と金額を入れる入力フォーム（UserForm）を足したい。シェルが使えないなら、どう進めるかの手順だけ教えて。
