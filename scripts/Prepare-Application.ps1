<#
.SYNOPSIS
    Prepares a built Access database for deployment by applying configuration.

.DESCRIPTION
    Opens the specified Access database file and applies post-build configuration
    from a JSON file. Supports:
      - Running VBA procedures (with optional parameters)
      - Setting DAO database properties (startup form, nav pane, bypass key, etc.)
      - Removing VBA modules by name pattern (e.g., test modules)
      - Removing VBA references by name (e.g., Rubberduck)

    This script runs AFTER the VCS build produces an .accdb and BEFORE any
    ACCDE compilation step, so that the compiled output reflects all changes.

    The JSON config format extends the upstream msaccess-vcs-build pattern
    (Procedures + DatabaseProperties) with RemoveModules and RemoveReferences.

.PARAMETER AccessFile
    Path to the Access database file (.accdb) to configure.

.PARAMETER ConfigFile
    Path to the JSON configuration file.

.EXAMPLE
    .\Prepare-Application.ps1 -AccessFile "tests\out\MyApp.accdb" -ConfigFile "Application-Config.json"

.NOTES
    Config JSON format (all keys optional; applied in the order below):
    {
      "RemoveModules": ["Tests_*"],          # name glob patterns (VBA components)
      "RemoveReferences": ["Rubberduck"],     # VBA reference names
      "Procedures": [                          # public procs run via Application.Run
        { "Name": "ChangeEnvironment", "Parameters": [3] }
      ],
      "DatabaseProperties": [                  # DAO database properties (created if absent)
        { "Name": "StartUpShowDBWindow", "Type": 1, "Value": false },
        { "Name": "StartUpForm", "Type": 10, "Value": "frmMainMenu" }
      ]
    }

    DatabaseProperties "Type" maps to the DAO DataTypeEnum used by CreateProperty:
      1  = Boolean
      3  = Integer
      4  = Long
      8  = DateTime
      10 = Text

    RemoveModules / RemoveReferences extend the upstream msaccess-vcs-build config
    (which supports only Procedures + DatabaseProperties).
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$AccessFile,

    [Parameter(Mandatory = $true)]
    [string]$ConfigFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -- Enums matching DAO property types ----------------------------------------

enum PropertyType {
    Boolean = 1
    Integer = 3
    Long = 4
    DateTime = 8
    Text = 10
}

# -- Helper functions ---------------------------------------------------------

function Set-DbProperty {
    param (
        [System.Object]$db,
        [string]$PropertyName,
        [int32]$PropertyType,
        $PropertyValue
    )

    try {
        $db.Properties[$PropertyName].Value = $PropertyValue
        Write-Host "  Updated property '$PropertyName' = '$PropertyValue'"
    }
    catch {
        $errorCode = $null
        if ($_.Exception.InnerException -and $_.Exception.InnerException.ErrorCode) {
            $errorCode = $_.Exception.InnerException.ErrorCode
        }
        elseif ($_.Exception.HResult) {
            $errorCode = $_.Exception.HResult
        }
        $errorMsg = $_.Exception.Message

        # Error -2146825018 or "Property not found" means we need to create it
        if ($errorCode -eq -2146825018 -or $errorMsg -like "*Property not found*") {
            Write-Host "  Property '$PropertyName' does not exist. Creating it."
            $db.Properties.Append($db.CreateProperty($PropertyName, $PropertyType, $PropertyValue))
        }
        else {
            Write-Error "  Unexpected error setting property '$PropertyName': $errorCode  $errorMsg"
        }
    }
}

function Invoke-Procedure {
    [CmdletBinding()]
    param (
        [System.Object]$access,
        [string]$ProcedureName,
        [object[]]$Arguments = @()
    )

    if (-not $access) {
        Write-Error "Access application object is null."
        return
    }
    if (-not $ProcedureName) {
        Write-Error "Procedure name is null or empty."
        return
    }

    # Cast numeric values to Int32 for COM/VBA compatibility
    # (ConvertFrom-Json produces Int64, but VBA Long is 32-bit)
    for ($i = 0; $i -lt $Arguments.Count; $i++) {
        if ($Arguments[$i] -is [long] -or $Arguments[$i] -is [int64]) {
            $Arguments[$i] = [int]$Arguments[$i]
        }
    }

    $ArgCount = $Arguments.Count
    Write-Host "  Invoke-Procedure: '$ProcedureName' with $ArgCount arg(s)"

    switch ($ArgCount) {
        0 { $null = $access.Run($ProcedureName) }
        1 { $null = $access.Run($ProcedureName, [ref] $Arguments[0]) }
        2 { $null = $access.Run($ProcedureName, [ref] $Arguments[0], [ref] $Arguments[1]) }
        3 { $null = $access.Run($ProcedureName, [ref] $Arguments[0], [ref] $Arguments[1], [ref] $Arguments[2]) }
        Default {
            Write-Error "Procedure '$ProcedureName' called with $ArgCount arguments, but Access.Run supports at most 3."
            return
        }
    }

    Write-Host "  Invoke-Procedure: '$ProcedureName' completed"
}

function Remove-VbaModules {
    param (
        [System.Object]$vbProject,
        [string[]]$Patterns
    )

    if (-not $Patterns -or $Patterns.Count -eq 0) {
        return
    }

    # Collect matching component names first to avoid modifying collection while iterating
    $componentsToRemove = @()
    foreach ($component in $vbProject.VBComponents) {
        foreach ($pattern in $Patterns) {
            if ($component.Name -like $pattern) {
                $componentsToRemove += $component.Name
                break
            }
        }
    }

    foreach ($name in $componentsToRemove) {
        try {
            $component = $vbProject.VBComponents.Item($name)
            $vbProject.VBComponents.Remove($component)
            Write-Host "  Removed module '$name'"
        }
        catch {
            Write-Host "  Warning: Could not remove module '$name': $($_.Exception.Message)"
        }
    }

    if ($componentsToRemove.Count -eq 0) {
        Write-Host "  No modules matched the removal patterns."
    }
}

function Remove-VbaReferences {
    param (
        [System.Object]$vbProject,
        [string[]]$ReferenceNames
    )

    if (-not $ReferenceNames -or $ReferenceNames.Count -eq 0) {
        return
    }

    foreach ($refName in $ReferenceNames) {
        $found = $false
        foreach ($ref in $vbProject.References) {
            if ($ref.Name -eq $refName) {
                try {
                    $vbProject.References.Remove($ref)
                    Write-Host "  Removed reference '$refName'"
                    $found = $true
                }
                catch {
                    Write-Host "  Warning: Could not remove reference '$refName': $($_.Exception.Message)"
                }
                break
            }
        }
        if (-not $found) {
            Write-Host "  Reference '$refName' not found (may already be absent)."
        }
    }
}

function SafeReleaseComObject($comObject) {
    if ($null -ne $comObject -and $comObject -is [System.__ComObject]) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($comObject)
    }
}

# -- Resolve paths ------------------------------------------------------------

if (-not ([System.IO.Path]::IsPathRooted($ConfigFile))) {
    $ConfigFile = Join-Path -Path (Get-Location) -ChildPath $ConfigFile.TrimStart('\', '/', '.')
}

if (-not (Test-Path -Path $ConfigFile)) {
    Write-Error "Config file not found: $ConfigFile"
    exit 1
}

if (-not ([System.IO.Path]::IsPathRooted($AccessFile))) {
    $AccessFile = Join-Path -Path (Get-Location) -ChildPath $AccessFile.TrimStart('\', '/', '.')
}

if (-not (Test-Path -Path $AccessFile)) {
    Write-Error "Access file not found: $AccessFile"
    exit 1
}

$config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json

Write-Host "=== Prepare Application ==="
Write-Host "Access file : $AccessFile"
Write-Host "Config file : $ConfigFile"
Write-Host ""

# -- Open Access and apply configuration --------------------------------------

[object]$access = $null
[object]$db = $null

try {
    $access = New-Object -ComObject Access.Application
    $access.OpenCurrentDatabase($AccessFile)

    # Step 1: Remove modules (before running procedures, in case test modules cause issues)
    if ($config.PSObject.Properties.Match('RemoveModules') -and $config.RemoveModules.Count -gt 0) {
        Write-Host "Removing VBA modules matching patterns: $($config.RemoveModules -join ', ')"
        $vbProject = $access.VBE.ActiveVBProject
        Remove-VbaModules -vbProject $vbProject -Patterns $config.RemoveModules
        Write-Host ""
    }

    # Step 2: Remove references (before running procedures, to avoid broken ref errors)
    if ($config.PSObject.Properties.Match('RemoveReferences') -and $config.RemoveReferences.Count -gt 0) {
        Write-Host "Removing VBA references: $($config.RemoveReferences -join ', ')"
        $vbProject = $access.VBE.ActiveVBProject
        Remove-VbaReferences -vbProject $vbProject -ReferenceNames $config.RemoveReferences
        Write-Host ""
    }

    # Step 3: Run procedures
    if ($config.PSObject.Properties.Match('Procedures') -and $config.Procedures.Count -gt 0) {
        Write-Host "Running procedures..."
        foreach ($procedure in $config.Procedures) {
            if (-not $procedure.Name) {
                Write-Error "Procedure name is missing in the configuration."
                continue
            }
            # Force into array — ConvertFrom-Json can unwrap single-element arrays to scalars,
            # and strict mode disallows .Count on scalars.
            [object[]]$Parameters = @()
            if ($procedure.PSObject.Properties['Parameters'] -and $null -ne $procedure.Parameters) {
                $Parameters = @($procedure.Parameters)
            }
            if ($Parameters.Count -gt 0) {
                Write-Host "  Running '$($procedure.Name)' with parameters: $($Parameters -join ', ')"
            }
            else {
                Write-Host "  Running '$($procedure.Name)'"
            }
            Invoke-Procedure -access $access -ProcedureName $procedure.Name -Arguments $Parameters
        }
        Write-Host ""
    }

    # Step 4: Set database properties
    if ($config.PSObject.Properties.Match('DatabaseProperties') -and $config.DatabaseProperties.Count -gt 0) {
        Write-Host "Setting database properties..."
        $db = $access.CurrentDb()

        foreach ($property in $config.DatabaseProperties) {
            $propertyName = $property.Name
            $propertyType = [PropertyType]::Parse([PropertyType], $property.Type)
            $propertyValue = $property.Value

            Write-Host "  Setting '$propertyName' (type $propertyType) = '$propertyValue'"
            Set-DbProperty -db $db -PropertyName $propertyName -PropertyType $propertyType -PropertyValue $propertyValue
        }
        Write-Host ""
    }

    Write-Host "Application preparation complete."
}
catch {
    $errorCode = $null
    if ($_.Exception.InnerException -and $_.Exception.InnerException.ErrorCode) {
        $errorCode = $_.Exception.InnerException.ErrorCode
    }
    elseif ($_.Exception.HResult) {
        $errorCode = $_.Exception.HResult
    }
    $errorMsg = $_.Exception.Message
    Write-Error "An error occurred during application preparation: $errorCode  $errorMsg"
    exit 1
}
finally {
    if ($db) {
        SafeReleaseComObject $db
        Remove-Variable -Name db -ErrorAction SilentlyContinue
    }
    if ($access) {
        $access.CloseCurrentDatabase()
        $access.Quit()
        SafeReleaseComObject $access
        Remove-Variable -Name access -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit 0
