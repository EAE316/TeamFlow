# TeamFlow 1.1.0 — SharePoint preview

Untested at the user's request. Compilation/packaging is not runtime or SharePoint
validation. Start with a disposable copy of a TeamFlow workbook.

## Connect
Open TeamFlow > workbook switch (החלפה) > SharePoint connection.
Paste the full HTTPS document address ending in .xlsx, or a SharePoint /:x:/r/ link.
For opaque sharing links: open in Excel > File > Info > Copy path.
Microsoft 365 desktop Excel must be installed, signed in, and allowed to coauthor.
The workbook needs existing TeamFlow headers and unique task IDs in column L.
Column T must be unused; TeamFlow stores per-task edit history there.
No Entra registration, password collection, cookie extraction, or OneDrive shortcut.

Excel opens visibly for sign-in, then minimizes. It remains open on TeamFlow exit
so Office can finish saving or display recovery prompts. Do not close Excel while
using this connection. Authentication/security dialogs are not suppressed.

## Included
- Read existing tasks from the live workbook in Excel.
- Edit existing tasks and mark them done, updating only their changed cells.
- Refresh from Excel every two seconds, preserving an open task draft.
- Reject a stale task version already visible in the local Excel session.
- Store write intent in the local excel-recovery folder before changing cells.
- Retain existing local-file mode and its task creation/deletion/settings features.

Cloud creation, deletion, bulk changes, row reordering, and list editing are disabled.
These require a separate coordination design; cloud preview does not yet replace all
local-mode features. Existing central history is read, but new cloud history is stored
per task in column T to avoid PCs appending to the same history row.

## Concurrency
Version checking is local, not a distributed lock or compare-and-swap. Two people
editing the same task/cells before Excel merges remote changes can still conflict.
Edit different tasks. Excel coauthoring does not provide database transactions.
"Connected to Excel / AutoSave on" does not confirm delivery to SharePoint or other PCs.
Check Excel's upload state. A write can partially succeed before a COM/network error;
a recovery record is retained and no automatic retry or rollback overwrites other edits.

## Launch and release
Excel libraries load only when needed. Cloud connection runs off the HTTP thread.
UPX is disabled. No before/after timing was measured.
Source backup: release-backups/source-before-cloud-*.
No application, browser, Excel, or SharePoint tests were run for this release.
The installed application and customer workbook were not upgraded or modified.

## Windows warning
This preview is unsigned unless a trusted production signing certificate is explicitly
supplied. Windows SmartScreen may still warn. No trust-root script is bundled or Windows
protection disabled. Store-distributed MSIX is the proposed free distribution route,
but publisher verification and Store submission have not been completed.

References:
- https://learn.microsoft.com/en-us/office/vba/excel/concepts/about-coauthoring-in-excel
- https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.autosaveon
- https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation
