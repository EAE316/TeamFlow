# TeamFlow 1.2.0 — preview

A fix release on top of 1.1.0. Test with a copy of your workbook first.

## Fixed
- **SharePoint connection dropping while Excel is busy.** If Excel is editing a cell or
  showing a dialog, TeamFlow now keeps the connection, shows "Excel עסוק · ממתין" and
  retries every two seconds. Other read errors are retried for 60 seconds before the
  connection stops. A save attempted while Excel is busy is refused without writing.
- **Task history moving to another task.** Deleting or reordering tasks in local-file mode
  now moves column T history, and any columns after S, with their task.
- **Formulas lost on local save.** A formula in a task row is kept until you change that
  field in TeamFlow. Calculated columns M–O always keep their formulas.
- **Failed connection attempts.** A failed SharePoint connection no longer replaces the
  current data source. You stay on the local file or the previous SharePoint file.
- Every connection stop is now written to the log with its reason.

## Still disconnects on purpose
Press the reconnect button after fixing the cause:
- The file becomes read-only, or AutoSave is turned off.
- A task row typed directly in Excel has no ID in column L, or an ID is duplicated.
- The task headers are changed.

## Known limits
- A formula edited by openpyxl has no stored result until Excel opens and recalculates the
  file, so TeamFlow can show that field as empty in local mode until then.
- Formulas outside the task rows that point at task cells are not adjusted when rows move.
- Rows already mismatched by 1.1.0 are not repaired.
- SharePoint mode edits existing tasks only. Creation, deletion, bulk changes, reordering
  and list editing stay disabled there.
- Version checking is local, not a distributed lock. Edit different tasks than others.
- "Connected to Excel" does not confirm upload to SharePoint. Check Excel's upload state.
- This preview is unsigned, so Windows may show an unknown-publisher warning.

## What was and was not tested
Automated tests cover the fixes using temporary workbooks and a simulated Excel.
The SharePoint changes have **not** been run against real Excel or SharePoint.

## Connect
Open TeamFlow > workbook switch (החלפה) > SharePoint connection.
Paste the full HTTPS document address ending in .xlsx, or a SharePoint /:x:/r/ link.
For opaque sharing links: open in Excel > File > Info > Copy path.
Microsoft 365 desktop Excel must be installed, signed in, and allowed to coauthor.
The workbook needs existing TeamFlow headers, unique task IDs in column L, and column T
free for TeamFlow's per-task history. Keep Excel running while using the connection.

## Going back
The 1.1.0 preview remains available on the releases page.
