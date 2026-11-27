; Inno Setup Script for TeamFlow - ניהול משימות צוות
#define MyAppName "TeamFlow"
#define MyAppHebrewName "ניהול משימות צוות"
#define MyAppVersion "1.4.2"
#define MyAppPublisher "Ibrahim Abd-Elnabi"
#define MyAppExeName "TeamTaskManagerWeb.exe"

[Setup]
AppId={{8B19E5FD-6591-4C5C-9A82-C86B49D8562B}
AppName={#MyAppName} - {#MyAppHebrewName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppComments=ניהול משימות צוות מתקדם בסנכרון Excel
AppCopyright=Copyright © 2026 Ibrahim Abd-Elnabi
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} - {#MyAppHebrewName}
VersionInfoVersion=1.4.2.0
VersionInfoTextVersion=1.4.2
VersionInfoOriginalFileName=TeamFlow_Setup.exe
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion=1.4.2.0
VersionInfoCopyright=Copyright © 2026 Ibrahim Abd-Elnabi
DefaultDirName={autopf}\TeamFlow
UsePreviousAppDir=yes
UsePreviousPrivileges=yes
UsePreviousTasks=yes
DisableDirPage=auto
CloseApplications=yes
CloseApplicationsFilter=TeamTaskManagerWeb.exe,*.dll,*.pyd
RestartApplications=no
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer_output
OutputBaseFilename=TeamFlow_Setup
SetupIconFile=app_icon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog commandline
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "hebrew"; MessagesFile: "compiler:Languages\Hebrew.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\TeamTaskManagerWeb\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "*.xlsx,*.lock,config.json,startup.log,uploads\*,excel-recovery\*"

; Remove only obsolete bundled runtimes, never customer data.
[InstallDelete]
Type: filesandordirs; Name: "{app}\_internal\numpy"
Type: filesandordirs; Name: "{app}\_internal\numpy.libs"
Type: filesandordirs; Name: "{app}\_internal\numpy-2.2.6.dist-info"
Type: filesandordirs; Name: "{app}\_internal\_tcl_data"
Type: filesandordirs; Name: "{app}\_internal\_tk_data"
Type: filesandordirs; Name: "{app}\_internal\tcl8"
Type: files; Name: "{app}\_internal\_tkinter.pyd"
Type: files; Name: "{app}\_internal\tcl86t.dll"
Type: files; Name: "{app}\_internal\tk86t.dll"

[Icons]
Name: "{autoprograms}\{#MyAppName} - {#MyAppHebrewName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName} - {#MyAppHebrewName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; Tasks: desktopicon; Check: ShouldCreateDesktopShortcut

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "הפעל את {#MyAppName} ({#MyAppHebrewName}) כעת"; Flags: nowait postinstall skipifsilent

[Code]
function ExistingDesktopShortcut: Boolean;
begin
  Result := FileExists(ExpandConstant('{userdesktop}\TeamFlow - ניהול משימות צוות.lnk')) or
    FileExists(ExpandConstant('{commondesktop}\TeamFlow - ניהול משימות צוות.lnk')) or
    FileExists(ExpandConstant('{userdesktop}\ניהול משימות צוות - TeamFlow.lnk')) or
    FileExists(ExpandConstant('{commondesktop}\ניהול משימות צוות - TeamFlow.lnk'));
end;

function ShouldCreateDesktopShortcut: Boolean;
begin
  Result := not ExistingDesktopShortcut;
end;

procedure UpdateExistingShortcut(const Filename: String);
begin
  if FileExists(Filename) then begin
    CreateShellLink(Filename, 'TeamFlow', ExpandConstant('{app}\{#MyAppExeName}'),
      '', ExpandConstant('{app}'), ExpandConstant('{app}\{#MyAppExeName}'), 0, SW_SHOWNORMAL);
    Log('Updated existing desktop shortcut: ' + Filename);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then begin
    UpdateExistingShortcut(ExpandConstant('{userdesktop}\TeamFlow - ניהול משימות צוות.lnk'));
    UpdateExistingShortcut(ExpandConstant('{commondesktop}\TeamFlow - ניהול משימות צוות.lnk'));
    UpdateExistingShortcut(ExpandConstant('{userdesktop}\ניהול משימות צוות - TeamFlow.lnk'));
    UpdateExistingShortcut(ExpandConstant('{commondesktop}\ניהול משימות צוות - TeamFlow.lnk'));
  end;
end;
