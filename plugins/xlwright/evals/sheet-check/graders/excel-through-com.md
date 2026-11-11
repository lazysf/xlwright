---
type: llm
---

PASS if the answer plans to open the workbook in Excel itself through PowerShell COM (for example with Invoke-ExcelScript.ps1), and to read the values back from Excel after saving.
FAIL if the answer plans to read or write the .xlsx directly with openpyxl, pandas, or another library that does not run Excel, or gives no way to check the result.
