# WinTuner VCIO - Winget Detection script
#
# Script parameters in `{parameter_name}`
# packageId - The package id of the application to be detected
# version   - The minimum version of the application to be detected

# --------------------------- Start parameters -------------------------------
$packageId = "{packageId}"
$version   = "{version}"
# --------------------------- End parameters ---------------------------------

# ------------------------------------ Start script -----------------------------------------
Start-Transcript -Path "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\$packageId-detection.log" -Force
Write-Host "Starting detection for packageId='$packageId' version='$version'"

#region Helper functions

function Get-WingetCmd {
    Write-Host "Resolving winget.exe path"

    # Prefer whatever is on PATH
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        Write-Host "Found winget on PATH at '$($cmd.Source)'"
        return $cmd.Source
    }

    # Fallbacks in case PATH lookup fails (rare on modern Windows, but cheap to try)
    $possible = @(
        "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\winget.exe",
        "$env:LocalAppData\Microsoft\WindowsApps\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\winget.exe"
    )

    foreach ($path in $possible) {
        if (Test-Path $path) {
            Write-Host "Found winget at '$path'"
            return $path
        }
    }

    Write-Host "winget.exe not found in PATH or known locations"
    return $null
}

function Compare-Versions {
    param(
        [Parameter(Mandatory)]
        [string] $VersionExpected,
        [Parameter(Mandatory)]
        [string] $VersionInstalled
    )

    # Returns:
    #  0  -> equal
    #  1  -> installed is higher than expected
    # -1  -> installed is lower than expected

    if ([string]::IsNullOrWhiteSpace($VersionExpected) -or
        [string]::IsNullOrWhiteSpace($VersionInstalled)) {
        return 0
    }

    try {
        $vExpected  = [version]$VersionExpected
        $vInstalled = [version]$VersionInstalled

        if ($vInstalled -gt $vExpected) { return 1 }
        if ($vInstalled -lt $vExpected) { return -1 }
        return 0
    } catch {
        # Fallback to string comparison when .NET version parsing cannot handle the format
        if ($VersionInstalled -eq $VersionExpected) { return 0 }
        if ($VersionInstalled -gt $VersionExpected) { return 1 }
        return -1
    }
}

function Get-WingetPackageFromJson {
    param(
        [Parameter(Mandatory)]
        [string] $WingetCmd,
        [Parameter(Mandatory)]
        [string] $PackageId
    )

    Write-Host "Querying winget (JSON) for packageId='$PackageId'"

    $raw = & $WingetCmd list --id $PackageId --exact --accept-source-agreements --output json 2>$null | Out-String
    if (-not $raw.Trim()) {
        Write-Host "No JSON output from 'winget list --id $PackageId --exact'"
        return $null
    }

    try {
        $data = $raw | ConvertFrom-Json
    } catch {
        Write-Host "Failed to parse winget JSON output: $($_.Exception.Message)"
        return $null
    }

    if (-not $data) {
        Write-Host "No entries returned in JSON for packageId='$PackageId'"
        return $null
    }

    # winget list can return a single object or an array
    if ($data -is [System.Array]) {
        $pkg = $data | Select-Object -First 1
    } else {
        $pkg = $data
    }

    Write-Host "JSON detection candidate: Name='$($pkg.Name)' Id='$($pkg.Id)' Version='$($pkg.Version)'"
    return $pkg
}

function Get-WingetPackageLegacy {
    param(
        [Parameter(Mandatory)]
        [string] $WingetCmd,
        [Parameter(Mandatory)]
        [string] $PackageId
    )

    Write-Host "Falling back to legacy text parsing for packageId='$PackageId'"

    $output = & $WingetCmd list --id $PackageId --exact --accept-source-agreements 2>$null
    if (-not $output) {
        Write-Host "No text output from 'winget list --id $PackageId --exact'"
        return $null
    }

    $lines = $output | Where-Object { $_.Trim() }
    if ($lines.Count -lt 2) {
        Write-Host "Not enough lines in winget output to parse (got $($lines.Count))"
        return $null
    }

    # Last non-empty line is the one with the package
    $header = $lines[0]
    $data   = $lines[-1]

    # Detect approximate column positions from the header
    $nameStart = 0
    $idStart   = $header.IndexOf("Id", [StringComparison]::OrdinalIgnoreCase)
    $verStart  = $header.IndexOf("Version", [StringComparison]::OrdinalIgnoreCase)

    if ($idStart -lt 0 -or $verStart -lt 0) {
        Write-Host "Could not identify Id/Version columns in winget output header"
        return $null
    }

    $name  = $data.Substring($nameStart, $idStart - $nameStart).Trim()
    $id    = $data.Substring($idStart,   $verStart - $idStart).Trim()
    $ver   = $data.Substring($verStart).Trim().Split(" ", 2)[0]

    $obj = [PSCustomObject]@{
        Name    = $name
        Id      = $id
        Version = $ver
        Source  = $null
    }

    Write-Host "Legacy detection candidate: Name='$($obj.Name)' Id='$($obj.Id)' Version='$($obj.Version)'"
    return $obj
}

#endregion Helper functions

#region Main

$wingetCmd = Get-WingetCmd
if (-not $wingetCmd) {
    Write-Host "winget.exe not available, exiting with code 10"
    Exit 10
}

# 1) Try strict Id lookup with JSON
$pkg = Get-WingetPackageFromJson -WingetCmd $wingetCmd -PackageId $packageId

# 2) If JSON-by-Id fails, try a broader JSON search by name/id
if (-not $pkg) {
    Write-Host "JSON lookup by Id failed, trying broader search by Id/Name"

    $rawSearch = & $wingetCmd list $packageId --accept-source-agreements --output json 2>$null | Out-String
    if ($rawSearch.Trim()) {
        try {
            $dataSearch = $rawSearch | ConvertFrom-Json
        } catch {
            $dataSearch = $null
            Write-Host "Failed to parse JSON for broader search: $($_.Exception.Message)"
        }

        if ($dataSearch) {
            if ($dataSearch -is [System.Array]) {
                $pkg = $dataSearch | Where-Object { $_.Id -eq $packageId -or $_.Name -eq $packageId } | Select-Object -First 1
            } else {
                if ($dataSearch.Id -eq $packageId -or $dataSearch.Name -eq $packageId) {
                    $pkg = $dataSearch
                }
            }

            if ($pkg) {
                Write-Host "Found package via broader JSON search: Name='$($pkg.Name)' Id='$($pkg.Id)' Version='$($pkg.Version)'"
            }
        }
    }
}

# 3) If JSON completely fails, fall back to legacy text parsing
if (-not $pkg) {
    $pkg = Get-WingetPackageLegacy -WingetCmd $wingetCmd -PackageId $packageId
}

if (-not $pkg) {
    Write-Host "Package '$packageId' not detected using winget, exiting with code 10"
    Exit 10
}

$installedVersion = [string]$pkg.Version
Write-Host "Detected installed version '$installedVersion' for packageId='$packageId'"

# If no version was provided, any installed version is success
if ([string]::IsNullOrWhiteSpace($version)) {
    Write-Host "No expected version specified, package is installed, exiting with code 0"
    Exit 0
}

$cmp = Compare-Versions -VersionExpected $version -VersionInstalled $installedVersion

if ($cmp -lt 0) {
    Write-Host "Installed version '$installedVersion' is lower than expected '$version', exiting with code 4"
    Exit 4
}

if ($cmp -eq 0) {
    Write-Host "Installed version '$installedVersion' equals expected '$version', exiting with code 0"
    Exit 0
}

# cmp > 0 -> installed is higher than expected
Write-Host "Installed version '$installedVersion' is higher than expected '$version', treating as compliant, exiting with code 0"
Exit 0

#endregion Main
