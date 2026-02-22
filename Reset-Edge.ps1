<#
.SYNOPSIS
    Performs a clean, repeatable reset of Microsoft Edge.

.DESCRIPTION
    Removes all Microsoft Edge user data, forces a bootstrap launch,
    and resets again to ensure a stable first run.

.AUTHOR
    sysadminsushi

.VERSION
    2.22.2026
#>

# Removes all Microsoft Edge user data and recreates the base folder so Edge can bootstrap cleanly.
function Reset-Edge {

    # Close all Microsoft Edge processes
    Get-Process -Name "msedge" -ErrorAction SilentlyContinue | Stop-Process -Force

    # Define the Microsoft Edge root directory path
    $edgeRootDirectoryPath = Join-Path $env:LOCALAPPDATA "Microsoft\Edge"

    # Remove all Microsoft Edge data
    if (Test-Path $edgeRootDirectoryPath) {
        Remove-Item -Path $edgeRootDirectoryPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Recreate the Microsoft Edge base directory
    if (-not (Test-Path $edgeRootDirectoryPath)) {
        New-Item -ItemType Directory -Path $edgeRootDirectoryPath | Out-Null
    }
}

# Performs a two‑stage reset to ensure Microsoft Edge bootstrap files are rebuilt cleanly and consistently.
function Reset-Edge-Twice {

    # Perform the initial reset
    Reset-Edge

    # Launch Microsoft Edge to force bootstrap creation
    Start-Process "msedge.exe" --no-first-run
    Start-Sleep -Seconds 3

    # Close Microsoft Edge after bootstrap initialization
    Get-Process -Name "msedge" -ErrorAction SilentlyContinue | Stop-Process -Force

    # Perform the second reset now that bootstrap files exist
    Reset-Edge

    return "Edge reset completed."
}

# Executes the reset sequence
Reset-Edge-Twice