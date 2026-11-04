---
name: xlwright
description: Rules for working on an Excel app folder with the xlwright tools (session_start, range_read, range_write, vba_get, vba_sync, vba_check, macro_run, test_run, workbook_save, session_end). Read it before the first session_start in a conversation.
---

# Working with xlwright

xlwright runs a dedicated hidden Excel for one app folder at a time. Follow these rules whenever you use its tools.

## The app folder

- An app folder holds at most one workbook directly in it, the main workbook (`.xlsm` or `.xlsx`). Put any other workbook, such as reference or test data, in a subfolder like `data/`.
- VBA source lives in the folder, not in the workbook:
  - `src/modules/*.bas` for standard modules
  - `src/classes/*.cls` for class modules
  - `src/workbook/<CodeName>.cls` for sheet and workbook modules
  - `tests/*.bas` and `tests/*.cls` for test modules
- Do not create subfolders under `src/` or `tests/`, and keep each file's extension matching its folder.
- Edit the source files, then run `vba_sync`, `vba_check`, `macro_run`, or `test_run` to bring them into the workbook. Do not edit `src/modules/XlWright.bas`; xlwright writes it.
- Name a file in `src/workbook/` after the CodeName that `vba_get` lists for the sheet, not after the sheet name.
- Do not nest app folders. Never put an app folder inside another app folder or above one.
- `.xlwright/` in the app folder is xlwright's runtime state. Keep it out of Git.

## The session

- A conversation has at most one session. Start it with `session_start` and pass the app folder.
- If `session_start` returns `cleanup`, xlwright ended a dedicated Excel left by an earlier session. Tell the user, with its `pid`.
- Call `session_end` before you report to the user, and decide `save` on purpose: `true` keeps the changes in the main workbook, `false` discards the changes since the last save.
- Do not write macros that change files outside the app folder.

## Checking and testing

- After you edit the VBA source, call `vba_check`. `compiled: false` is a result, not a failure: fix the module and line in `error`, then call it again.
- A test is a `Sub` without arguments, not `Private`, whose name starts with `Test`, in `tests/*.bas`. Assert with `XlWrightAssert.AreEqual expected, actual[, message]`, `XlWrightAssert.IsTrue condition[, message]`, and `XlWrightAssert.Fail message`.
- `test_run` saves the main workbook first, including the changes made before the call. Then it runs the tests in a copy of the app folder, so their changes do not reach the main workbook or the app folder.
- A failing test is a result, not a failure. Fix the source or the test, then call `test_run` again.
- In a test, reach the data files through `ThisWorkbook.Path`. Do not write to the main workbook's folder through a full path, and do not change files outside the app folder.
- If a test opens a workbook outside the app folder, close it in the same test. `test_run` closes only the workbooks in the copy, and a workbook left open makes `session_end` leave the dedicated Excel running (`FOREIGN_WORKBOOK`).

## Failures

- Read `message` in the error and follow it. It says what to do next.
- After `STATE_UNKNOWN`, call `session_start` again only when both `data.excel_ended` and `data.restartable` are `true`. Otherwise follow `message`.
- Do not resend a call whose result is unknown on your own.
- If an operation that changes something was too large to answer in full, the response gives a summary and the path of the full result under `.xlwright/results/`. Do not call it again to get the rest.
