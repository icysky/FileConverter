# One-click build script for File Converter installer.
# Usage:
#   .\build-installer.ps1                     -> Release MSI
#   .\build-installer.ps1 -Configuration Debug -> Debug MSI
#   .\build-installer.ps1 -Restore            -> force NuGet restore first
# NOTE: keep this file ASCII only (PowerShell 5.1 reads BOM-less files as ANSI).

param(
    [ValidateSet("Release", "Debug")]
    [string]$Configuration = "Release",
    [switch]$Restore
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$solution = Join-Path $repoRoot "FileConverter.sln"

Write-Host "=== File Converter installer build ===" -ForegroundColor Cyan
Write-Host "Solution     : $solution"
Write-Host "Configuration: $Configuration (x64)"

# 1. Locate MSBuild via vswhere (works for any VS edition), fallback to known path.
$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    $installPath = & $vswhere -latest -requires Microsoft.Component.MSBuild -property installationPath
    $msbuild = Join-Path $installPath "MSBuild\Current\Bin\MSBuild.exe"
}
else {
    $msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
}

if (-not (Test-Path $msbuild)) {
    Write-Error "MSBuild.exe not found. Install Visual Studio 2022 with the .NET desktop workload."
}

Write-Host "MSBuild      : $msbuild"
Write-Host ""

# 2. Optional NuGet restore.
if ($Restore) {
    Write-Host "--- Restoring NuGet packages ---" -ForegroundColor Cyan
    & $msbuild $solution /t:Restore /p:Configuration=$Configuration /p:Platform=x64 /v:q /nologo
    if ($LASTEXITCODE -ne 0) { Write-Error "NuGet restore failed (exit code $LASTEXITCODE)." }
}

# 3. Build the whole solution. Building the solution (not the wixproj alone) is required:
#    the app post-build events copy ffmpeg.exe / Ghostscript into the output folder and
#    need $(SolutionDir); the WiX project then harvests that output as MSI payload.
Write-Host "--- Building $Configuration x64 (app + shell extension + WiX installer) ---" -ForegroundColor Cyan
& $msbuild $solution /p:Configuration=$Configuration /p:Platform=x64 /v:m /nologo
if ($LASTEXITCODE -ne 0) { Write-Error "Build failed (exit code $LASTEXITCODE)." }

# 4. Report artifact.
$msi = Join-Path $repoRoot "Installer\bin\x64\$Configuration\FileConverter-setup.msi"
if (-not (Test-Path $msi)) { Write-Error "MSI not found at expected path: $msi" }

$sizeMb = [math]::Round((Get-Item $msi).Length / 1MB, 1)
[xml]$versionXml = Get-Content (Join-Path $repoRoot "version.xml")
$latest = $versionXml.Version.Latest
$version = "$($latest.major).$($latest.minor).$($latest.patch)"
Write-Host ""
Write-Host "=== Build succeeded ===" -ForegroundColor Green
Write-Host "MSI    : $msi"
Write-Host "Version: $version   Size: ${sizeMb} MB"
