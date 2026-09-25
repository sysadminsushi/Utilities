<#
.SYNOPSIS
    Resets the Microsoft OneDrive client (non-AVD).
.DESCRIPTION
    Finds OneDrive.exe (per-user or per-machine) and runs /reset.
.EXAMPLE
    Reset-OneDriveEngine
.NOTES
    Author:  sysadminsushi
    Version: 9.25.2026
#>
function Get-OneDriveExecutablePath {
    [CmdletBinding()]
    param()
    $candidateExecutablePaths = @(
        (Join-Path $env:LOCALAPPDATA "Microsoft\OneDrive\OneDrive.exe"),
        (Join-Path $env:ProgramFiles "Microsoft OneDrive\OneDrive.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft OneDrive\OneDrive.exe")
    )
    foreach ($candidateExecutablePath in $candidateExecutablePaths) {
        if ($candidateExecutablePath -and (Test-Path -LiteralPath $candidateExecutablePath)) {
            return $candidateExecutablePath
        }
    }
    return $null
}

function Reset-OneDriveEngine {
    [CmdletBinding()]
    param()
    $oneDriveExecutablePath = Get-OneDriveExecutablePath
    if (-not $oneDriveExecutablePath) {
        Write-Warning "Could not find OneDrive.exe in expected locations."
        return
    }
    Write-Output "Resetting OneDrive: $oneDriveExecutablePath"
    Start-Process -FilePath $oneDriveExecutablePath -ArgumentList "/reset"
}

Reset-OneDriveEngine
