#Requires -Version 5.1
<#
.SYNOPSIS
    Builds a single .accdb from a msaccess-vcs source folder (headless).

.DESCRIPTION
    Drives the locally installed MSAccess VCS add-in to build one exported source
    folder into an .accdb, written into TargetDir. Access runs invisibly and the
    caller's working databases are never touched.

    The add-in writes its output one directory level up from the source folder it is
    given, so this script builds from a scratch copy of the source placed *inside*
    TargetDir; the resulting .accdb therefore lands directly in TargetDir. TargetDir
    is created if missing but never wiped here -- clearing it (for single-database
    builds) or preserving it (for layered multi-database builds) is the caller's job.

    On success the full path of the built .accdb is written to the success stream as
    the script's only pipeline output, so a caller can capture it:

        $accdb = & .\Build-Accdb.ps1 -SourceDir ... -TargetDir ... | Select-Object -Last 1

.PARAMETER SourceDir
    Path to the VCS source folder to build. Relative paths resolve from the CWD.

.PARAMETER TargetDir
    Directory that will receive the built .accdb. Relative paths resolve from the CWD.

.PARAMETER VcsAddInPath
    Full path to the installed "Version Control.accda" add-in.

.PARAMETER TimeoutSeconds
    Per-pass seconds to wait for the add-in's progress forms to close. Default 120.

.EXAMPLE
    $accdb = & .\Build-Accdb.ps1 -SourceDir "AMP_Hobby_Database.accde" -TargetDir "build" | Select-Object -Last 1
#>
param(
    [Parameter(Mandatory = $true)][string]$SourceDir,
    [Parameter(Mandatory = $true)][string]$TargetDir,
    [string]$VcsAddInPath = (Join-Path $env:AppData "MSAccessVCS\Version Control.accda"),
    [int]$TimeoutSeconds = 120
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -- Resolve paths ------------------------------------------------------------

if (-not [System.IO.Path]::IsPathRooted($SourceDir)) { $SourceDir = Join-Path (Get-Location) $SourceDir }
if (-not [System.IO.Path]::IsPathRooted($TargetDir)) { $TargetDir = Join-Path (Get-Location) $TargetDir }
$SourceDir = [System.IO.Path]::GetFullPath($SourceDir)
$TargetDir = [System.IO.Path]::GetFullPath($TargetDir)

# Strip extension -- used with Application.Run("path.FunctionName").
$AddInProcessPath = [System.IO.Path]::ChangeExtension($VcsAddInPath, "").TrimEnd('.')

if (-not (Test-Path $VcsAddInPath)) {
    Write-Error "msaccess-vcs add-in not found: $VcsAddInPath"
    exit 1
}
if (-not (Test-Path $SourceDir)) {
    Write-Error "Source directory not found: $SourceDir"
    exit 1
}

New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null

# Seed host DB for the add-in API (it requires an open database). Kept in temp, out
# of TargetDir, so it never pollutes the output or collides with a parallel build.
$SeedDbPath = Join-Path ([System.IO.Path]::GetTempPath()) ("VcsBuildSeed_{0}.accdb" -f ([Guid]::NewGuid().ToString('N')))

# Scratch source copy inside TargetDir; the add-in writes one level up => into
# TargetDir. Prefixed so it never collides with the source/compiled name, and
# suffixed with the source leaf so layered multi-database builds don't clash.
$SourceLeaf = [System.IO.Path]::GetFileName($SourceDir.TrimEnd('\'))
$BuildSrcDir = Join-Path $TargetDir ("_vcs-build-src-" + $SourceLeaf)
if (Test-Path $BuildSrcDir) { Remove-Item $BuildSrcDir -Recurse -Force }
Copy-Item -Path $SourceDir -Destination $BuildSrcDir -Recurse -Force
Write-Host "Building '$SourceLeaf' -> $TargetDir"

$access = $null
$accessProcessId = $null
$BuiltFilePath = $null
$BuiltFileName = $null

try {
    # Snapshot existing Access PIDs. Invisible Access has MainWindowHandle 0, so
    # identify the process this COM instance starts by PID diff instead.
    $existingAccessPids = @(
        Get-Process -Name MSACCESS -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty Id
    )

    $access = New-Object -ComObject Access.Application
    $access.Visible = $false

    try {
        for ($i = 0; $i -lt 10 -and $null -eq $accessProcessId; $i++) {
            $accessProcessId = Get-Process -Name MSACCESS -ErrorAction SilentlyContinue |
                Where-Object { $existingAccessPids -notcontains $_.Id } |
                Select-Object -First 1 -ExpandProperty Id
            if ($null -ne $accessProcessId) { break }
            Start-Sleep -Milliseconds 200
        }
    }
    catch { $accessProcessId = $null }

    if (Test-Path $SeedDbPath) { Remove-Item $SeedDbPath -Force }
    $access.NewCurrentDatabase($SeedDbPath)

    $access.Run("$AddInProcessPath.SetInteractionMode", [ref] 1)
    # Supported public API dispatcher -> clsVersionControl.Build (not the ribbon click).
    $null = $access.Run("$AddInProcessPath.API", [ref] "Build", [ref] "$BuildSrcDir")

    # Progress forms close when the build finishes; poll twice (upstream pattern).
    Write-Host "Waiting for VCS build to complete" -NoNewline
    foreach ($pass in 1..2) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while (($access.Forms.Count -gt 0) -and ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds)) {
            Start-Sleep -Seconds 2
            Write-Host "." -NoNewline
        }
        $sw.Stop()
        Start-Sleep -Seconds 3
    }
    Write-Host " done"

    $BuiltFileName = $access.CurrentProject.Name
    $BuiltFilePath = $access.CurrentProject.FullName

    if ([string]::IsNullOrEmpty($BuiltFilePath) -or ($BuiltFilePath -ieq $SeedDbPath)) {
        Write-Error "VCS build failed -- Access still has the seed database open, not the built one."
        exit 1
    }

    Write-Host "Built: $BuiltFileName"
}
finally {
    if ($null -ne $access) {
        # Do NOT Quit() headlessly: a full build can leave frmVCSMain open, whose
        # unload prompt ("Cancel Current Operation?") blocks a graceful Quit with no
        # way to answer it. The build is complete (CurrentProject read above), so
        # release the RCW and terminate the tracked process directly instead.
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
        $access = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    if ($null -ne $accessProcessId) {
        Get-Process -Id $accessProcessId -ErrorAction SilentlyContinue |
            Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

# Wait for the process to exit so its file locks are released.
$sw = [System.Diagnostics.Stopwatch]::StartNew()
while (($null -ne $accessProcessId) -and (Get-Process -Id $accessProcessId -ErrorAction SilentlyContinue) -and ($sw.Elapsed.TotalSeconds -lt 30)) {
    Start-Sleep -Milliseconds 500
}
$sw.Stop()

$TargetAccdbPath = Join-Path $TargetDir $BuiltFileName
$resolvedBuilt = [System.IO.Path]::GetFullPath($BuiltFilePath)
$resolvedTarget = [System.IO.Path]::GetFullPath($TargetAccdbPath)
if ($resolvedBuilt -ine $resolvedTarget) {
    Move-Item -Path $BuiltFilePath -Destination $TargetAccdbPath -Force
}

if (Test-Path $BuildSrcDir) { Remove-Item $BuildSrcDir -Recurse -Force }
if (Test-Path $SeedDbPath) { Remove-Item $SeedDbPath -Force }

Write-Host "ACCDB ready: $TargetAccdbPath"

# Sole pipeline output: the built database path (for the caller to capture).
Write-Output $TargetAccdbPath
exit 0
