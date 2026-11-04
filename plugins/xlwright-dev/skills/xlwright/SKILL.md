---
name: xlwright
description: Rules for working on an Excel app folder with the xlwright tools (session_start, range_read, range_write, vba_sync, macro_run, workbook_save, session_end). Read it before the first session_start in a conversation.
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
- Edit the source files, then run `vba_sync` or `macro_run` to bring them into the workbook. Do not edit `src/modules/XlWright.bas`; xlwright writes it.
- Do not nest app folders. Never put an app folder inside another app folder or above one.
- `.xlwright/` in the app folder is xlwright's runtime state. Keep it out of Git.

## The session

- A conversation has at most one session. Start it with `session_start` and pass the app folder.
- If `session_start` returns `cleanup`, xlwright ended a dedicated Excel left by an earlier session. Tell the user, with its `pid`.
- Call `session_end` before you report to the user, and decide `save` on purpose: `true` keeps the changes in the main workbook, `false` discards the changes since the last save.
- Do not write macros that change files outside the app folder.

## Failures

- Read `message` in the error and follow it. It says what to do next.
- After `STATE_UNKNOWN`, call `session_start` again only when both `data.excel_ended` and `data.restartable` are `true`. Otherwise follow `message`.
- Do not resend a call whose result is unknown on your own.
- If an operation that changes something was too large to answer in full, the response gives a summary and the path of the full result under `.xlwright/results/`. Do not call it again to get the rest.
