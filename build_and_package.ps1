# Build and package a clean TeamFlow release, excluding local customer data.
param(
    [string]$Python = 'python',
    [string]$Iscc = '',
    [string]$CertificateThumbprint = '',
    [string]$TimestampServer = 'http://timestamp.digicert.com',
    [switch]$SkipTests
)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

function Invoke-CheckedNative([string]$Executable, [string[]]$Arguments) {
    # Windows PowerShell 5 treats redirected native stderr as ErrorRecords.
    # Python and PyInstaller use stderr for normal progress; trust their exit code.
    $previousPreference = $ErrorActionPreference
    $nativeExitCode = 0
    try {
        $ErrorActionPreference = 'Continue'
        & $Executable @Arguments 2>&1 | ForEach-Object { Write-Host "$_" }
        $nativeExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousPreference
    }
    if ($nativeExitCode -ne 0) { throw "$Executable failed with exit code $nativeExitCode" }
}

if (-not $SkipTests) {
    Invoke-CheckedNative $Python @('tests/run_backend.py')
    Invoke-CheckedNative 'node' @('--check', 'web/app.js')
}

Invoke-CheckedNative $Python @('-m', 'PyInstaller', 'TeamTaskManagerWeb.spec', '--noconfirm')
$stage = Join-Path $PSScriptRoot 'dist\TeamTaskManagerWeb'
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Install-TeamTaskManager.ps1') -Destination $stage -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'TeamTaskManagerWeb\Run-TeamFlow.bat') -Destination $stage -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'TeamTaskManagerWeb\Install.bat') -Destination $stage -Force

Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'RELEASE_NOTES.md') -Destination $stage -Force

# A self-signed certificate is not a clean-PC SmartScreen solution.
$cert = $null
if ($CertificateThumbprint) {
    if ($CertificateThumbprint -notmatch '^[a-fA-F0-9]{40}$') { throw 'Invalid certificate thumbprint.' }
    $cert = Get-Item ("Cert:\CurrentUser\My\" + $CertificateThumbprint) -ErrorAction Stop
    if (-not $cert.HasPrivateKey -or $cert.Subject -eq $cert.Issuer) {
        throw 'A production code-signing certificate with a private key is required.'
    }
}
function Sign-ReleaseFile([string]$File) {
    if ($cert) {
        $signature = Set-AuthenticodeSignature -FilePath $File -Certificate $cert -HashAlgorithm SHA256 -TimestampServer $TimestampServer
        if ($signature.Status -ne 'Valid' -or -not $signature.TimeStamperCertificate) {
            throw "Signing or timestamping failed: $File ($($signature.Status))"
        }
        Write-Host "Signed and timestamped: $File"
    } else {
        Write-Warning "Unsigned preview: Windows may show an unknown-publisher/SmartScreen warning."
    }
}
Sign-ReleaseFile (Join-Path $stage 'TeamTaskManagerWeb.exe')

if (-not $Iscc) {
    $compiler = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($compiler) { $Iscc = $compiler.Source }
    else { $Iscc = Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe' }
}
if (-not (Test-Path -LiteralPath $Iscc)) { throw "Inno Setup compiler not found: $Iscc" }
Invoke-CheckedNative $Iscc @((Join-Path $PSScriptRoot 'TeamFlow.iss'))
$installer = Join-Path $PSScriptRoot 'installer_output\TeamFlow_Setup.exe'
Sign-ReleaseFile $installer
Copy-Item -LiteralPath $installer -Destination (Join-Path $PSScriptRoot 'TeamFlow_Setup.exe') -Force
Compress-Archive -Path "$stage\*" -DestinationPath (Join-Path $PSScriptRoot 'TeamFlow-Standalone-Windows.zip') -Force
Compress-Archive -LiteralPath (Join-Path $PSScriptRoot 'TeamFlow_Setup.exe') -DestinationPath (Join-Path $PSScriptRoot 'TeamFlow_Setup.zip') -Force
Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $PSScriptRoot 'TeamFlow_Setup.exe'), (Join-Path $PSScriptRoot 'TeamFlow-Standalone-Windows.zip') | Format-Table -AutoSize
Write-Host 'Build and packaging completed.'
