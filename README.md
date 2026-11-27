# TeamFlow

Windows desktop task manager with a Hebrew interface and Excel workbook storage.

## Download
Get the installer EXE from [Releases](../../releases).
Version 1.4.2 is a **preview** with a new modern interface (light and dark themes), dates
shown as dd/mm/yyyy, and full task editing in SharePoint mode. Its SharePoint code was tested
with desktop Excel on local copies, not with real SharePoint or two PCs editing at once.
Earlier previews stay available on the same page.
The installer is unsigned and Windows may display an unknown-publisher warning.

## SharePoint preview
Connect through the installed Microsoft 365 desktop Excel application using a full
SharePoint .xlsx document URL. Excel handles Office sign-in and coauthoring.
No TeamFlow Entra registration or OneDrive shortcut is required by this design.

Everything is done in the app, also in SharePoint mode: creating, editing and deleting
tasks, bulk actions, manual card order, comments and the team lists. Writes are cell-level,
rows never move, and new tasks get random IDs so two PCs cannot clash. Two people editing
the *same task* within seconds of each other is not locked: the last save wins. AutoSave
enabled is not cloud delivery confirmation. Start with a disposable workbook copy.
See [release notes](RELEASE_NOTES.md).

## Build
Windows, Python 3.12, and Inno Setup 6 are required.

    python -m pip install -r requirements-build.txt
    .\build_and_package.ps1 -Python python -SkipTests

This source distribution does not include private workbooks, local configuration,
certificates, logs, backups, or the previous private test fixtures. Use -SkipTests.
Tests were run before release from the private fixtures, including tests that drive real
desktop Excel on local copies. No real SharePoint tests were run for 1.4.2.
