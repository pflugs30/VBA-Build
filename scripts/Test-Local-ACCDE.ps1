#Requires -Version 5.1
<#
.SYNOPSIS
    Tests the ACCDE build pipeline locally without GitHub Actions.

.DESCRIPTION
    Replicates what the AccessCodeLib/msaccess-vcs-build action does in two stages:
      1. Builds an .accdb from VCS source using the locally installed msaccess-vcs-addin.
      2. Compiles that .accdb to .accde via Access SysCmd(603).

    The script derives the repo root from its own location, so it can be run from anywhere.

.PARAMETER SourceDir
    Path to the VCS source folder (the .accde source folder).
    Relative paths are resolved from the repo root.
    Default: tests/AccessDatabase.accde

.PARAMETER TargetDir
    Output directory for the built files.
    Relative paths are resolved from the repo root.
    Default: tests/out

.PARAMETER VcsAddInPath
    Full path to the installed Version Control.accda addin file.

.EXAMPLE
    .\scripts\Test-Local-ACCDE.ps1

.EXAMPLE
    .\scripts\Test-Local-ACCDE.ps1 -SourceDir "tests/AccessDatabase.accde" -TargetDir "tests/out"
#>
param(
    [string]$SourceDir = "tests/AccessDatabase.accde",
    [string]$TargetDir = "tests/out",
    [string]$VcsAddInPath = "C:\Users\pflug\AppData\Roaming\MSAccessVCS\Version Control.accda"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -- Resolve paths ------------------------------------------------------------

$RepoRoot = Split-Path $PSScriptRoot -Parent

if (-not [System.IO.Path]::IsPathRooted($SourceDir)) {
    $SourceDir = Join-Path $RepoRoot $SourceDir
}
if (-not [System.IO.Path]::IsPathRooted($TargetDir)) {
    $TargetDir = Join-Path $RepoRoot $TargetDir
}

$SourceDir = [System.IO.Path]::GetFullPath($SourceDir)
$TargetDir = [System.IO.Path]::GetFullPath($TargetDir)

# Strip extension -- used with Application.Run("path.FunctionName")
$AddInProcessPath = [System.IO.Path]::ChangeExtension($VcsAddInPath, "").TrimEnd('.')

# -- Pre-flight checks --------------------------------------------------------

if (-not (Test-Path $VcsAddInPath)) {
    Write-Error "msaccess-vcs-addin not found: $VcsAddInPath"
    exit 1
}
if (-not (Test-Path $SourceDir)) {
    Write-Error "Source directory not found: $SourceDir"
    exit 1
}

Write-Host "Repo root : $RepoRoot"
Write-Host "Source    : $SourceDir"
Write-Host "Target    : $TargetDir"
Write-Host "Add-in    : $VcsAddInPath"
Write-Host ""

# -- Step 1: Build .accdb from VCS source -------------------------------------

Write-Host "=== Step 1: Build .accdb from VCS source ==="

# Ensure output directory exists upfront.
New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null

# The addin derives the output directory from the fixture's ExportFolder path,
# so the .accdb will be built at the repo root regardless of CWD.
# We move it to TargetDir afterwards to keep the repo root clean.
$TempDbPath = Join-Path $RepoRoot "VcsBuildTempApp.accdb"

$access = $null
$BuiltFilePath = $null
$BuiltFileName = $null

try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $true

    if (Test-Path $TempDbPath) { Remove-Item $TempDbPath -Force }
    $access.NewCurrentDatabase($TempDbPath)
    Write-Host "Seed database created: $TempDbPath"

    Write-Host "Calling VCS build with source: $SourceDir"
    $access.Run("$AddInProcessPath.SetInteractionMode", [ref] 1)
    $null = $access.Run("$AddInProcessPath.HandleRibbonCommand", [ref] "btnBuild", [ref] "$SourceDir")

    # The addin shows progress forms while importing; forms closing = build done.
    # Poll twice with a pause between (matches upstream Build-Accdb.ps1 logic).
    Write-Host "Waiting for VCS build to complete" -NoNewline
    foreach ($pass in 1..2) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while (($access.Forms.Count -gt 0) -and ($sw.Elapsed.TotalSeconds -lt 120)) {
            Start-Sleep -Seconds 2
            Write-Host "." -NoNewline
        }
        $sw.Stop()
        Start-Sleep -Seconds 3
    }
    Write-Host " done"

    $BuiltFileName = $access.CurrentProject.Name
    $BuiltFilePath = $access.CurrentProject.FullName

    if ([string]::IsNullOrEmpty($BuiltFilePath) -or ($BuiltFileName -ieq "VcsBuildTempApp.accdb")) {
        Write-Error "VCS build failed -- Access still has the seed database open, not the built one."
        exit 1
    }

    Write-Host "Built: $BuiltFileName  ($BuiltFilePath)"

}
finally {
    if ($null -ne $access) {
        try { $access.Quit(2) } catch {}
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
        $access = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$TargetAccdbPath = Join-Path $TargetDir $BuiltFileName
Move-Item -Path $BuiltFilePath -Destination $TargetAccdbPath -Force
Write-Host "Moved .accdb to: $TargetAccdbPath"

if (Test-Path $TempDbPath) {
    Remove-Item $TempDbPath -Force
    Write-Host "Removed seed database."
}

# Wait for Access to fully terminate before starting Step 2 (prevents COM state leaks).
Write-Host "Waiting for Access to shut down..." -NoNewline
$sw = [System.Diagnostics.Stopwatch]::StartNew()
while ((Get-Process MSACCESS -ErrorAction SilentlyContinue) -and ($sw.Elapsed.TotalSeconds -lt 30)) {
    Start-Sleep -Milliseconds 500
}
$sw.Stop()
Write-Host " done"

Write-Host ""

# -- Step 2: Compile .accdb to .accde -----------------------------------------

Write-Host "=== Step 2: Compile .accdb to .accde ==="

$AccdeDestPath = [System.IO.Path]::ChangeExtension($TargetAccdbPath, "accde")
Write-Host "Source : $TargetAccdbPath"
Write-Host "Dest   : $AccdeDestPath"

if (Test-Path $AccdeDestPath) { Remove-Item $AccdeDestPath -Force }

# Run compile in a subprocess to avoid COM apartment state from Step 1.
# This mirrors CI where compile is a separate action step / process.
$compileScript = Join-Path $env:TEMP "vba-build-compile-$PID.ps1"
@"
param([string]`$SourceFile, [string]`$DestFile)
`$access = New-Object -ComObject Access.Application
`$accessType = `$access.GetType()
`$null = `$accessType.InvokeMember('SysCmd', 'InvokeMethod', `$null, `$access, @(603, `$SourceFile, `$DestFile))
`$ok = `$false
for (`$i = 0; `$i -lt 30; `$i++) {
    if (Test-Path `$DestFile) { `$ok = `$true; break }
    Start-Sleep -Seconds 1
}
`$access.Quit(2)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$access)
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
if (-not `$ok) { Write-Error 'accde file was not created.'; exit 1 }
"@ | Set-Content -Path $compileScript -Encoding UTF8

try {
    Write-Host "Compiling (subprocess)..."
    pwsh -NoProfile -ExecutionPolicy Bypass -File $compileScript `
        -SourceFile $TargetAccdbPath -DestFile $AccdeDestPath
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Compile step failed (exit $LASTEXITCODE)."
        exit 1
    }
}
finally {
    Remove-Item $compileScript -Force -ErrorAction SilentlyContinue
}

$CompileSuccess = Test-Path $AccdeDestPath

# -- Result ------------------------------------------------------------------

if ($CompileSuccess) {
    $size = (Get-Item $AccdeDestPath).Length
    Write-Host ""
    Write-Host "SUCCESS: ACCDE created: $AccdeDestPath ($size bytes)"
    exit 0
}
else {
    Write-Host ""
    Write-Error "FAILED: ACCDE was not created at: $AccdeDestPath"
    exit 1
}
