<#
.SYNOPSIS
  Performs a full hard reset of classic Outlook (Outlook.exe) for the current user in AVD/VDI environments.

.DESCRIPTION
  This script stops classic Outlook for the current user session, removes all classic Outlook cache data,
  detects the installed Office version, removes all Outlook profiles under the user's registry hive,
  and optionally recreates an empty Profiles key if -RecreateProfileKey is specified.

  New Outlook (Olk) and the MSIX/Store "Outlook for Windows" app are not affected.

.AUTHOR
  sysadminsushi

.VERSION
  2.25.2026
#>

# Retrieves the current user's session ID for process‑scoped termination
function Get-CurrentUserSessionId {
    try {
        $currentProcess = Get-Process -Id $PID
        return $currentProcess.SessionId
    } catch {
        return $null
    }
}

# Stops Outlook.exe only for the current user's session
function Stop-CurrentUserOutlookProcesses {
    $currentUserSessionId = Get-CurrentUserSessionId
    $processNamesToStop = @('Outlook')

    foreach ($processName in $processNamesToStop) {
        try {
            $matchingProcesses = Get-Process -Name $processName -ErrorAction SilentlyContinue
            if ($matchingProcesses) {
                if ($currentUserSessionId -ne $null) {
                    $matchingProcesses |
                        Where-Object { $_.SessionId -eq $currentUserSessionId } |
                        Stop-Process -Force -ErrorAction SilentlyContinue
                } else {
                    $matchingProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
                }
            }
        } catch {}
    }

    Start-Sleep -Milliseconds 800
}

# Removes all classic Outlook cache directories for the current user
function Remove-ClassicOutlookCaches {
    $pathsToClearForClassicOutlook = @(
        (Join-Path $env:LocalAppData 'Microsoft\Outlook\*'),
        (Join-Path $env:AppData 'Microsoft\Outlook\*')
    )

    foreach ($cachePath in $pathsToClearForClassicOutlook) {
        try {
            if (Test-Path $cachePath) {
                Remove-Item -Path $cachePath -Recurse -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
}

# Detects the installed Office version by locating the Outlook Profiles root
function Get-InstalledOfficeOutlookProfileRoot {
    $possibleOfficeVersions = @('16.0', '15.0', '14.0')
    foreach ($officeVersion in $possibleOfficeVersions) {
        $candidatePath = "HKCU:\Software\Microsoft\Office\$officeVersion\Outlook\Profiles"
        if (Test-Path $candidatePath) {
            return $candidatePath
        }
    }
    return $null
}

# Removes all Outlook profiles under the detected Profiles key
function Remove-AllOutlookProfiles {
    param(
        [string]$outlookProfilesRootPath
    )

    try {
        if (Test-Path $outlookProfilesRootPath) {
            Remove-Item -Path $outlookProfilesRootPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

# Recreates an empty Outlook Profiles key if requested
function New-EmptyOutlookProfilesKey {
    param(
        [string]$outlookProfilesRootPath
    )

    try {
        New-Item -Path $outlookProfilesRootPath -Force | Out-Null
    } catch {}
}

# Main entry point for performing the classic Outlook hard reset
function Invoke-AVDClassicOutlookHardReset {
    param(
        [switch]$RecreateProfileKey
    )

    Stop-CurrentUserOutlookProcesses
    Remove-ClassicOutlookCaches

    $outlookProfilesRootPath = Get-InstalledOfficeOutlookProfileRoot
    if ($outlookProfilesRootPath) {
        Remove-AllOutlookProfiles -outlookProfilesRootPath $outlookProfilesRootPath

        if ($RecreateProfileKey) {
            New-EmptyOutlookProfilesKey -outlookProfilesRootPath $outlookProfilesRootPath
        }
    }
}

# Executes the reset sequence
Invoke-AVDClassicOutlookHardReset