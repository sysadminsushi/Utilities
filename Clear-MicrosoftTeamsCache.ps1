<#
.SYNOPSIS
    Clears the Microsoft Teams (New Teams) cache for the current user.

.DESCRIPTION
    Stops the Microsoft Teams process, verifies the cache directory exists,
    and removes all cached data to ensure a clean restart.

.AUTHOR
    sysadminsushi

.VERSION
    2.22.2026
#>

# Clears the Microsoft Teams cache by stopping the process and removing cached data from the local cache directory.
function Clear-MicrosoftTeamsCache {

    # Stop Microsoft Teams if it is currently running
    Get-Process -Name "ms-teams" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    # Define the Microsoft Teams cache directory path
    $teamsCacheDirectoryPath = Join-Path $env:LOCALAPPDATA "Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams"

    # Remove all cached Teams data if the directory exists
    if (Test-Path $teamsCacheDirectoryPath) {
        try {
            Remove-Item -Path "$teamsCacheDirectoryPath\*" -Recurse -Force -ErrorAction SilentlyContinue
        } catch {}
    }
}

# Executes the Teams cache clearing process
Clear-MicrosoftTeamsCache