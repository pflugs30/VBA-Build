#Requires -Version 5.1
<#
.SYNOPSIS
    Builds (and optionally compiles) one Access database from VCS source.

.DESCRIPTION
    Orchestrates the reusable build stages for a single database:
      1. Build-Accdb.ps1          -- VCS source folder -> .accdb
      2. Prepare-Application.ps1  -- (optional) apply Application-Config.json
      3. Compile-Accdb.ps1        -- .accdb -> .accde (unless -SkipAccdeCompile)

    Stage 3 runs in a fresh PowerShell subprocess to isolate its COM apartment
    from the preceding in-process automation (matching the CI step boundary).

    By default TargetDir is wiped before building. Pass -NoClean to preserve it,
    which is how several databases are layered into one shared output folder (their
    references resolve at runtime by co-location -- see handbook/buildScripts.md).

.PARAMETER SourceDir
    VCS source folder to build. Relative paths resolve from the CWD.

.PARAMETER TargetDir
    Output directory for the built .accdb/.accde. Relative paths resolve from the CWD.

.PARAMETER VcsAddInPath
    Full path to the installed "Version Control.accda" add-in.

.PARAMETER AppConfigFile
    Optional Application-Config.json applied after the build. Schema is documented
    in handbook/buildScripts.md and in Prepare-Application.ps1.

.PARAMETER SkipAccdeCompile
    Stop after building/preparing the .accdb; skip ACCDE compilation.

.PARAMETER OpenAccdbAfterBuild
    With -SkipAccdeCompile, open the built .accdb when done (for review/testing).

.PARAMETER NoClean
    Do not wipe TargetDir first (used when layering multiple databases into it).

.EXAMPLE
    .\Build-Application.ps1 -SourceDir "AMP_Hobby_Database.accde" -TargetDir "build" -AppConfigFile "Application-Config.json"

.EXAMPLE
    .\Build-Application.ps1 -SourceDir "AMP_Hobby_Database.accde" -TargetDir "build" -SkipAccdeCompile -OpenAccdbAfterBuild
#>
param(
    [Parameter(Mandatory = $true)][string]$SourceDir,
    [Parameter(Mandatory = $true)][string]$TargetDir,
    [string]$VcsAddInPath = (Join-Path $env:AppData "MSAccessVCS\Version Control.accda"),
    [string]$AppConfigFile = "",
    [switch]$SkipAccdeCompile,
    [switch]$OpenAccdbAfterBuild,
    [switch]$NoClean
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptDir = $PSScriptRoot
$BuildAccdbScript = Join-Path $ScriptDir "Build-Accdb.ps1"
$PrepareScript = Join-Path $ScriptDir "Prepare-Application.ps1"
$CompileScript = Join-Path $ScriptDir "Compile-Accdb.ps1"

foreach ($s in @($BuildAccdbScript, $CompileScript)) {
    if (-not (Test-Path $s)) { Write-Error "Required script not found: $s"; exit 1 }
}

# -- Resolve paths ------------------------------------------------------------

if (-not [System.IO.Path]::IsPathRooted($SourceDir)) { $SourceDir = Join-Path (Get-Location) $SourceDir }
if (-not [System.IO.Path]::IsPathRooted($TargetDir)) { $TargetDir = Join-Path (Get-Location) $TargetDir }
$SourceDir = [System.IO.Path]::GetFullPath($SourceDir)
$TargetDir = [System.IO.Path]::GetFullPath($TargetDir)

if (-not [string]::IsNullOrEmpty($AppConfigFile)) {
    if (-not [System.IO.Path]::IsPathRooted($AppConfigFile)) { $AppConfigFile = Join-Path (Get-Location) $AppConfigFile }
    $AppConfigFile = [System.IO.Path]::GetFullPath($AppConfigFile)
    if (-not (Test-Path $AppConfigFile)) { Write-Error "App config file not found: $AppConfigFile"; exit 1 }
}

Write-Host "Source    : $SourceDir"
Write-Host "Target    : $TargetDir"
Write-Host "Add-in    : $VcsAddInPath"
if ($AppConfigFile) { Write-Host "App config: $AppConfigFile" }
Write-Host "Skip ACCDE: $SkipAccdeCompile"
Write-Host ""

# -- Clean target (single-database default) -----------------------------------

if (-not $NoClean -and (Test-Path $TargetDir)) {
    Write-Host "Cleaning target folder: $TargetDir"
    Remove-Item $TargetDir -Recurse -Force
}
New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null

# -- Stage 1: Build .accdb ----------------------------------------------------

Write-Host "=== Stage 1: Build .accdb from source ==="
$AccdbPath = & $BuildAccdbScript -SourceDir $SourceDir -TargetDir $TargetDir -VcsAddInPath $VcsAddInPath |
    Select-Object -Last 1
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($AccdbPath) -or -not (Test-Path $AccdbPath)) {
    Write-Error "Build stage failed."
    exit 1
}
Write-Host ""

# -- Stage 2: Prepare application (optional) ----------------------------------

if ($AppConfigFile) {
    Write-Host "=== Stage 2: Prepare application ==="
    & $PrepareScript -AccessFile $AccdbPath -ConfigFile $AppConfigFile
    if ($LASTEXITCODE -ne 0) { Write-Error "Prepare stage failed."; exit 1 }
    Write-Host ""
}

# -- Stop here if ACCDE compilation is not requested --------------------------

if ($SkipAccdeCompile) {
    $accdbSize = (Get-Item $AccdbPath).Length
    Write-Host "SUCCESS: ACCDB created: $AccdbPath ($accdbSize bytes)"
    if ($OpenAccdbAfterBuild) {
        Write-Host "Opening ACCDB for review/testing..."
        Start-Process -FilePath $AccdbPath | Out-Null
    }
    exit 0
}

# -- Stage 3: Compile .accdb to .accde (fresh subprocess) ---------------------

Write-Host "=== Stage 3: Compile .accdb to .accde ==="
$AccdePath = [System.IO.Path]::ChangeExtension($AccdbPath, "accde")

$psExe = (Get-Command pwsh -ErrorAction SilentlyContinue).Path
if ([string]::IsNullOrWhiteSpace($psExe)) { $psExe = (Get-Command powershell -ErrorAction Stop).Path }

& $psExe -NoProfile -ExecutionPolicy Bypass -File $CompileScript -SourceFile $AccdbPath -DestFile $AccdePath
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $AccdePath)) {
    Write-Error "Compile stage failed (exit $LASTEXITCODE)."
    exit 1
}

$size = (Get-Item $AccdePath).Length
Write-Host ""
Write-Host "SUCCESS: ACCDE created: $AccdePath ($size bytes)"
exit 0
