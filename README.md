# TeamFlow

Windows desktop task manager with a Hebrew interface and Excel workbook storage.

## Download
Get the installer EXE from [Releases](../../releases).
Version 1.1.0 is an **untested preview**, not a production-validated release.
The installer is unsigned and Windows may display an unknown-publisher warning.

## SharePoint preview
Connect through the installed Microsoft 365 desktop Excel application using a full
SharePoint .xlsx document URL. Excel handles Office sign-in and coauthoring.
No TeamFlow Entra registration or OneDrive shortcut is required by this design.

The preview supports **editing existing tasks only** in cloud mode. Creation,
deletion, bulk operations, row reordering, and list editing are disabled there.
Local-file mode retains those features. Same-task simultaneous edits are not
protected by a distributed lock. AutoSave enabled is not cloud delivery confirmation.
Start with a disposable workbook copy. See [release notes](RELEASE_NOTES.md).

## Build
Windows, Python 3.12, and Inno Setup 6 are required.

    python -m pip install -r requirements-build.txt
    .\build_and_package.ps1 -Python python -SkipTests

This source distribution does not include private workbooks, local configuration,
certificates, logs, backups, or the previous private test fixtures. Use -SkipTests.
No runtime or SharePoint tests were run for the preview release.
