<#
.SYNOPSIS
    Azure Virtual Desktop–safe OneDrive engine reset that discovers the correct executable location and restarts as the standard user.

.DESCRIPTION
    Detects OneDrive across common per-user and per-machine installation paths, performs a OneDrive engine reset (/reset),
    waits briefly for cleanup, and restarts OneDrive using explorer.exe so it launches under the standard user context.
    Designed specifically for multi-session AVD so the scope stays within the current user session.

.AUTHOR
    sysadminsushi

.VERSION
    2.28.2026
#>

# Returns the best-matched OneDrive.exe path for the current device and user
function Get-AVDOneDriveExecutablePath {
    [CmdletBinding()]
    param()

    $oneDriveUserAppDataExecutablePath     = Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'
    $oneDriveProgramFilesExecutablePath    = 'C:\Program Files\Microsoft OneDrive\OneDrive.exe'
    $oneDriveProgramFilesX86ExecutablePath = 'C:\Program Files (x86)\Microsoft OneDrive\OneDrive.exe'

    $candidateExecutablePathsInPriorityOrder = @(
        $oneDriveUserAppDataExecutablePath,
        $oneDriveProgramFilesExecutablePath,
        $oneDriveProgramFilesX86ExecutablePath
    )

    foreach ($candidateExecutablePath in $candidateExecutablePathsInPriorityOrder) {
        if (Test-Path -LiteralPath $candidateExecutablePath) {
            return $candidateExecutablePath
        }
    }

    return $null
}

# Launches OneDrive as the standard, interactive user using explorer.exe
function Start-AVDOneDriveAsStandardUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OneDriveExecutablePath
    )

    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$OneDriveExecutablePath`""
}

# Invokes the OneDrive engine reset with path discovery and user-safe restart for AVD
function Reset-AVDOneDriveEngine {
    [CmdletBinding()]
    param()

    $discoveredOneDriveExecutablePath = Get-AVDOneDriveExecutablePath

    if ($null -ne $discoveredOneDriveExecutablePath) {
        Start-Process -FilePath $discoveredOneDriveExecutablePath -ArgumentList '/reset'
        Start-Sleep -Seconds 5
        Start-AVDOneDriveAsStandardUser -OneDriveExecutablePath $discoveredOneDriveExecutablePath
    } else {
        Write-Warning 'Could not find OneDrive.exe in expected locations. Please verify installation.'
    }
}