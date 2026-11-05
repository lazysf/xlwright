---
name: xlwright
description: Rules for working on an Excel app folder with the xlwright tools (session_start, range_read, range_write, vba_get, vba_edit, vba_extract, vba_sync, vba_check, macro_run, test_run, workbook_create, workbook_get, workbook_edit, workbook_save, session_end). Read it before the first session_start in a conversation.
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
- If the folder has no workbook yet, `session_start` starts a session without one (`workbook` is `null`). Call `workbook_create` with a file name such as `Sales.xlsm` before the other tools.
- VBA needs a `.xlsm` main workbook. To turn a `.xlsx` into one, call `workbook_edit` with `action` `convert_to_xlsm`. The `.xlsx` stays next to it, and the next `session_start` stops with `MULTIPLE_WORKBOOKS` until you remove the `.xlsx` with Git.

## An existing workbook

- Each sync replaces every standard and class module with `src/`, so a module that is only in the workbook is lost. Before the first `vba_sync`, `vba_check`, `macro_run`, or `test_run` on a workbook whose VBA is not in `src/` yet, call `vba_extract`.
- `vba_extract` does not write UserForms; it lists them in `skipped`, and the syncs leave them in the workbook. Tell the user about them.
- `vba_extract` stops with `SOURCE_EXISTS` when `src/` already has files. It does not overwrite them.
- References are kept in the workbook, not in `src/`. List them with `vba_get` and add or remove one with `vba_edit`, passing `guid`, `major`, and `minor` as `vba_get` gives them.

## Dialogs

- In the app's code, show a dialog only through `XlWright.Message` (in place of `MsgBox`), `XlWright.InputText`, `XlWright.OpenFile`, `XlWright.SaveFile`, or `XlWright.Folder`. Outside xlwright they show the standard dialog. Any other dialog stops the macro, and xlwright ends the dedicated Excel.
- Rewrite the dialogs in source taken from an existing workbook in the same way.
- For `macro_run`, pass the answers in `dialogs`, in the order the macro asks. A path is relative to the app folder. `dialogs.unused` in the result counts the answers the macro did not ask for.
- In a test, queue the answers with `XlWright.Answer kind, value` before the code asks. `False` cancels. Each test starts with no answers.
- If the answers run out, or the next one is of another kind, the macro stops with `MACRO_ERROR` whose `source` is `XlWright`, and `description` says what did not match. Make the answers match the dialogs the code shows, in order.

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
- `test_run` compiles the copy before the tests. With `COMPILE_ERROR` no test ran and the session stays open: write what the test calls, or fix the module and line in `data.error`, then call `test_run` again.
- In a test, reach the data files through `ThisWorkbook.Path`. Do not write to the main workbook's folder through a full path, and do not change files outside the app folder.
- If a test opens a workbook outside the app folder, close it in the same test. `test_run` closes only the workbooks in the copy, and a workbook left open makes `session_end` leave the dedicated Excel running (`FOREIGN_WORKBOOK`).

## Failures

- Read `message` in the error and follow it. It says what to do next.
- After `STATE_UNKNOWN`, call `session_start` again only when both `data.excel_ended` and `data.restartable` are `true`. Otherwise follow `message`.
- Do not resend a call whose result is unknown on your own.
- If an operation that changes something was too large to answer in full, the response gives a summary and the path of the full result under `.xlwright/results/`. Do not call it again to get the rest.
