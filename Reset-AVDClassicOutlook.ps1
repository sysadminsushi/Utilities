<#
.SYNOPSIS
    Performs a full hard reset of classic Outlook (Outlook.exe) for the current user in AVD/VDI environments.

.DESCRIPTION
    Stops classic Outlook for the current user session, removes classic Outlook user cache data,
    detects the installed Office version, removes all Outlook profiles under the user's registry hive,
    and optionally recreates an empty Profiles key if -RecreateProfileKey is specified.

    New Outlook (Olk) and the MSIX/Store "Outlook for Windows" app are not affected.

.AUTHOR
    sysadminsushi

.VERSION
    2.28.2026
#>

# Retrieves the current user's session ID for process-scoped termination in multi-session hosts
function Get-CurrentUserSessionId {
    [CmdletBinding()]
    param()

    try {
        $currentProcessForSessionDiscovery = Get-Process -Id $PID
        return $currentProcessForSessionDiscovery.SessionId
    } catch {
        return $null
    }
}

# Stops classic Outlook processes running only in the current user's session
function Stop-CurrentUserOutlookProcesses {
    [CmdletBinding()]
    param()

    $currentUserSessionId = Get-CurrentUserSessionId
    $processNamesTargetList = @('Outlook')

    foreach ($processNameToTerminate in $processNamesTargetList) {
        try {
            $allMatchingProcesses = Get-Process -Name $processNameToTerminate -ErrorAction SilentlyContinue
            if ($allMatchingProcesses) {
                if ($currentUserSessionId -ne $null) {
                    $allMatchingProcesses |
                        Where-Object { $_.SessionId -eq $currentUserSessionId } |
                        Stop-Process -Force -ErrorAction SilentlyContinue
                } else {
                    $allMatchingProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
                }
            }
        } catch {}
    }

    Start-Sleep -Milliseconds 800
}

# Removes classic Outlook cache directories for the current user profile
function Remove-ClassicOutlookCaches {
    [CmdletBinding()]
    param()

    $classicOutlookCachePathsToClear = @(
        (Join-Path $env:LocalAppData 'Microsoft\Outlook\*'),
        (Join-Path $env:AppData 'Microsoft\Outlook\*')
    )

    foreach ($classicOutlookCachePath in $classicOutlookCachePathsToClear) {
        try {
            if (Test-Path $classicOutlookCachePath) {
                Remove-Item -Path $classicOutlookCachePath -Recurse -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
}

# Detects the installed Office version by discovering the Outlook Profiles registry root under HKCU
function Get-InstalledOfficeOutlookProfileRoot {
    [CmdletBinding()]
    param()

    $officeVersionCandidates = @('16.0', '15.0', '14.0')
    foreach ($officeVersionCandidate in $officeVersionCandidates) {
        $candidateProfilesRootPath = "HKCU:\Software\Microsoft\Office\$officeVersionCandidate\Outlook\Profiles"
        if (Test-Path $candidateProfilesRootPath) {
            return $candidateProfilesRootPath
        }
    }
    return $null
}

# Removes all Outlook profiles beneath the detected Profiles registry root
function Remove-AllOutlookProfiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutlookProfilesRootPath
    )

    try {
        if (Test-Path $OutlookProfilesRootPath) {
            Remove-Item -Path $OutlookProfilesRootPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

# Recreates an empty Outlook Profiles registry key if requested
function New-EmptyOutlookProfilesKey {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutlookProfilesRootPath
    )

    try {
        New-Item -Path $OutlookProfilesRootPath -Force | Out-Null
    } catch {}
}

# Orchestrates a complete hard reset of classic Outlook for the current user in AVD/VDI
function Reset-AVDClassicOutlook {
    [CmdletBinding()]
    param(
        [switch]$RecreateProfileKey
    )

    Stop-CurrentUserOutlookProcesses
    Remove-ClassicOutlookCaches

    $detectedOutlookProfilesRootPath = Get-InstalledOfficeOutlookProfileRoot
    if ($detectedOutlookProfilesRootPath) {
        Remove-AllOutlookProfiles -OutlookProfilesRootPath $detectedOutlookProfilesRootPath

        if ($RecreateProfileKey) {
            New-EmptyOutlookProfilesKey -OutlookProfilesRootPath $detectedOutlookProfilesRootPath
        }
    }
}

Reset-AVDClassicOutlook