---
type: llm
---

PASS if the steps the answer gives include, after exporting, reading VBProject.Protection back from the workbooks: 1 (locked) for the exported workbook and 0 (unlocked) for the working workbook. The answer may only describe these steps without running them.
FAIL if the answer treats the output as locked without reading the protection back, or locks the working workbook itself.
