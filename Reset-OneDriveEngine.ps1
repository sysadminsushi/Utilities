<#
.SYNOPSIS
    Performs a reset and restart of the Microsoft OneDrive client engine (non-AVD).

.DESCRIPTION
    Locates the OneDrive executable across common installation paths (per-user and per-machine),
    executes the OneDrive engine reset (/reset) to remediate sync issues, waits briefly for cleanup,
    and restarts OneDrive via explorer.exe so it runs under the standard user context even if the
    script was launched elevated.

.AUTHOR
    sysadminsushi

.VERSION
    2.28.2026

.EXAMPLE
    Reset-OneDriveEngine
#>

# Returns the best-matched OneDrive.exe path
function Get-OneDriveExecutablePath {
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
function Start-OneDriveAsStandardUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OneDriveExecutablePath
    )

    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$OneDriveExecutablePath`""
}

# Performs a OneDrive engine reset and restarts it as the standard user
function Reset-OneDriveEngine {
    [CmdletBinding()]
    param()

    $discoveredOneDriveExecutablePath = Get-OneDriveExecutablePath

    if ($null -ne $discoveredOneDriveExecutablePath) {
        Write-Host "OneDrive located. Starting engine reset..." -ForegroundColor Cyan
        Start-Process -FilePath $discoveredOneDriveExecutablePath -ArgumentList '/reset'
        Start-Sleep -Seconds 5
        Write-Host "Restarting OneDrive as a standard user..." -ForegroundColor Green
        Start-OneDriveAsStandardUser -OneDriveExecutablePath $discoveredOneDriveExecutablePath
    } else {
        Write-Warning 'Could not find OneDrive.exe in expected locations. Please verify installation.'
    }
}

Reset-OneDriveEngine