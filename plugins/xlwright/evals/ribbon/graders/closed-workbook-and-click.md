---
type: llm
---

PASS if the answer embeds the customUI XML while the workbook is closed (not through Excel COM), puts the callback in a standard VBA module, and then checks the button by actually clicking it in a visible Excel (for example with Invoke-RibbonButton.ps1).
FAIL if the answer defines the ribbon through Excel COM, edits the workbook with openpyxl, or treats the button as working without clicking it.
