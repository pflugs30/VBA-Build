<#
.SYNOPSIS
    Compiles an Access .accdb database to .accde (compiled, locked VBA).

.DESCRIPTION
    Uses Access SysCmd(603) to compile an .accdb file into an .accde file.
    This is the same mechanism used by Access's "Make ACCDE" feature.

    Runs in a fresh COM instance to avoid apartment state issues from
    prior Access automation steps.

.PARAMETER SourceFile
    Path to the .accdb file to compile.

.PARAMETER DestFile
    Path for the output .accde file. If omitted, uses the same name/location
    with .accde extension.

.PARAMETER TimeoutSeconds
    Maximum seconds to wait for the .accde file to appear. Default: 60.

.EXAMPLE
    .\Compile-Accdb.ps1 -SourceFile "tests\out\MyApp.accdb"

.EXAMPLE
    .\Compile-Accdb.ps1 -SourceFile "tests\out\MyApp.accdb" -DestFile "dist\MyApp.accde"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceFile,

    [string]$DestFile = "",

    [int]$TimeoutSeconds = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -- Resolve paths ------------------------------------------------------------

if (-not ([System.IO.Path]::IsPathRooted($SourceFile))) {
    $SourceFile = Join-Path -Path (Get-Location) -ChildPath $SourceFile.TrimStart('\', '/', '.')
}
$SourceFile = [System.IO.Path]::GetFullPath($SourceFile)

if (-not (Test-Path $SourceFile)) {
    Write-Error "Source file not found: $SourceFile"
    exit 1
}

if ([string]::IsNullOrEmpty($DestFile)) {
    $DestFile = [System.IO.Path]::ChangeExtension($SourceFile, "accde")
}
if (-not ([System.IO.Path]::IsPathRooted($DestFile))) {
    $DestFile = Join-Path -Path (Get-Location) -ChildPath $DestFile.TrimStart('\', '/', '.')
}
$DestFile = [System.IO.Path]::GetFullPath($DestFile)

if (Test-Path $DestFile) {
    Remove-Item $DestFile -Force
}

Write-Host "=== Compile ACCDB to ACCDE ==="
Write-Host "Source : $SourceFile"
Write-Host "Dest   : $DestFile"

# -- Compile ------------------------------------------------------------------

$access = $null

try {
    $access = New-Object -ComObject Access.Application
    $accessType = $access.GetType()

    # SysCmd 603 = acSysCmdMakeACCDEFile
    $null = $accessType.InvokeMember(
        'SysCmd', 'InvokeMethod', $null, $access, @(603, $SourceFile, $DestFile)
    )

    # Wait for the .accde file to appear
    $ok = $false
    for ($i = 0; $i -lt $TimeoutSeconds; $i++) {
        if (Test-Path $DestFile) {
            $ok = $true
            break
        }
        Start-Sleep -Seconds 1
    }

    if (-not $ok) {
        Write-Error "ACCDE file was not created within $TimeoutSeconds seconds."
        exit 1
    }

    $size = (Get-Item $DestFile).Length
    Write-Host "ACCDE created: $DestFile ($size bytes)"
}
finally {
    if ($null -ne $access) {
        try { $access.Quit(2) } catch {}
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
        $access = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
