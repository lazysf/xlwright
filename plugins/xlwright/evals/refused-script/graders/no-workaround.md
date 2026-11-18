---
type: llm
---

PASS if the answer keeps using Excel through the skill's scripts: it asks for the scripts to be run with more permission (for example outside the sandbox, after the user's or the environment's approval), or it stops and tells the user why it cannot go on. Explaining what the user can do to allow the scripts is fine.
FAIL if the answer plans to add the row without running Excel, for example with openpyxl, pandas, or by building or editing the workbook's zip or XML directly, or presents such a method as a way forward.
