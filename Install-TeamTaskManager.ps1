$ErrorActionPreference = "Stop"

# A registered installation must be updated by the installer that owns it.
# Creating another portable installation would split its data and shortcuts.
$appId = '{8B19E5FD-6591-4C5C-9A82-C86B49D8562B}_is1'
foreach ($hive in @('HKCU:', 'HKLM:')) {
    if (Test-Path -LiteralPath ($hive + '\Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $appId)) {
        throw 'TeamFlow is already installed. Run TeamFlow_Setup.exe to update the existing installation.'
    }
}

if (Test-Path (Join-Path $PSScriptRoot "TeamTaskManagerWeb.exe")) {
    $source = $PSScriptRoot
} elseif (Test-Path (Join-Path $PSScriptRoot "TeamTaskManagerWeb\TeamTaskManagerWeb.exe")) {
    $source = Join-Path $PSScriptRoot "TeamTaskManagerWeb"
} else {
    throw "Missing TeamTaskManagerWeb.exe in $PSScriptRoot"
}

$target = Join-Path $env:LOCALAPPDATA "TeamTaskManager"
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "ניהול משימות צוות - TeamFlow.lnk"
$canonicalShortcut = Join-Path $desktop 'TeamFlow - ניהול משימות צוות.lnk'
if (Test-Path -LiteralPath $canonicalShortcut) { $shortcutPath = $canonicalShortcut }

Write-Host "Installing TeamFlow..." -ForegroundColor Cyan

$installedExe = Join-Path $target 'TeamTaskManagerWeb.exe'
$running = Get-CimInstance Win32_Process -Filter "Name = 'TeamTaskManagerWeb.exe'" |
    Where-Object { $_.ExecutablePath -eq $installedExe }
if ($running) { throw 'Close TeamFlow using its Exit button before updating it.' }

New-Item -ItemType Directory -Path $target -Force | Out-Null
# Update application files while preserving customer workbooks, uploads and config.
foreach ($item in Get-ChildItem -LiteralPath $source -Force) {
    if ($item.Name -in @('config.json', 'startup.log', 'uploads', '.instance.json') -or $item.Extension -in @('.xlsx', '.lock')) {
        continue
    }
    Copy-Item -LiteralPath $item.FullName -Destination $target -Recurse -Force
}

# Old optional runtimes can shadow the new bundled modules if left in place.
$resolvedTarget = [IO.Path]::GetFullPath($target).TrimEnd('\')
foreach ($runtime in @('numpy', 'numpy.libs', 'numpy-2.2.6.dist-info', '_tcl_data', '_tk_data', 'tcl8', '_tkinter.pyd', 'tcl86t.dll', 'tk86t.dll')) {
    $obsoletePath = [IO.Path]::GetFullPath((Join-Path $target ('_internal\' + $runtime)))
    if (-not $obsoletePath.StartsWith($resolvedTarget + '\_internal\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Unexpected runtime cleanup path.'
    }
    if (Test-Path -LiteralPath $obsoletePath) { Remove-Item -LiteralPath $obsoletePath -Recurse -Force }
}

$exe = Join-Path $target "TeamTaskManagerWeb.exe"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exe
$shortcut.WorkingDirectory = $target
$shortcut.IconLocation = "$exe,0"
$shortcut.Description = "ניהול משימות צוות - TeamFlow"
$shortcut.Save()

Write-Host "Successfully installed to: $target" -ForegroundColor Green
Write-Host "Desktop shortcut created: $shortcutPath" -ForegroundColor Green
