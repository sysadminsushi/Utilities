<#
.SYNOPSIS
    Clears the new Microsoft Teams cache for the current user.

.DESCRIPTION
    Stops ms-teams, then deletes LocalCache under the MSTeams Store package.
    Does not clear classic Teams (%APPDATA%\Microsoft\Teams).

.NOTES
    Author:  sysadminsushi
    Version: 9.24.2026
#>
function Clear-MicrosoftTeamsCache {
    Get-Process -Name "ms-teams" -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3

    $teamsCacheDirectoryPath = Join-Path $env:LOCALAPPDATA "Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams"

    if (-not (Test-Path $teamsCacheDirectoryPath)) {
        Write-Output "Teams cache path not found: $teamsCacheDirectoryPath"
        return
    }

    Remove-Item -Path (Join-Path $teamsCacheDirectoryPath "*") -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "Cleared Teams cache: $teamsCacheDirectoryPath"
}

Clear-MicrosoftTeamsCache
