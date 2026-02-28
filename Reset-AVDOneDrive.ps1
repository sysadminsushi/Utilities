<#
.SYNOPSIS
    Resets OneDrive for the current user in Azure Virtual Desktop (multi-session) safely.

.DESCRIPTION
    Stops OneDrive processes only for the current user's session, invokes the OneDrive engine reset,
    clears user-scoped identity and account registry caches, and removes local settings to ensure a
    clean sign-in prompt on next launch. Designed for multi-session AVD so other users are not affected.

.AUTHOR
    sysadminsushi

.VERSION
    2.28.2026
#>

# Stops OneDrive processes running in the current user's AVD session only
function Stop-AvdUserOneDriveProcess {
    [CmdletBinding()]
    param()

    process {
        $currentSessionId = (Get-Process -Id $PID).SessionId
        $currentSessionOneDriveProcesses = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue |
            Where-Object { $_.SessionId -eq $currentSessionId }

        if ($currentSessionOneDriveProcesses) {
            Write-Host "Stopping OneDrive process for the current AVD session..." -ForegroundColor Yellow
            $currentSessionOneDriveProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }
    }
}

# Runs the OneDrive engine reset switch for the current user if the executable is present
function Invoke-AvdOneDriveEngineReset {
    [CmdletBinding()]
    param()

    process {
        $oneDriveExecutablePath = Join-Path $env:LocalAppData "Microsoft\OneDrive\OneDrive.exe"
        if (Test-Path -LiteralPath $oneDriveExecutablePath) {
            Write-Host "Triggering OneDrive engine reset..." -ForegroundColor Cyan
            Start-Process -FilePath $oneDriveExecutablePath -ArgumentList "/reset" -NoNewWindow -Wait
            Start-Sleep -Seconds 2
        } else {
            Write-Host "OneDrive executable not found at: $oneDriveExecutablePath" -ForegroundColor DarkYellow
        }
    }
}

# Removes user-scoped registry caches that hold OneDrive account and identity information
function Clear-AvdOneDriveUserRegistryCaches {
    [CmdletBinding()]
    param()

    process {
        $registryPathsToPurge = @(
            "HKCU:\Software\Microsoft\OneDrive\Accounts",
            "HKCU:\Software\Microsoft\IdentityCRL\UserExtendedProperties",
            "HKCU:\Software\Microsoft\OneDrive\Settings"
        )

        foreach ($registryPath in $registryPathsToPurge) {
            if (Test-Path -LiteralPath $registryPath) {
                Write-Host "Removing registry cache: $registryPath" -ForegroundColor Gray
                Remove-Item -Path $registryPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# Deletes the user-scoped local OneDrive settings folder
function Remove-AvdOneDriveLocalSettingsFolder {
    [CmdletBinding()]
    param()

    process {
        $oneDriveSettingsFolderPath = Join-Path $env:LocalAppData "Microsoft\OneDrive\settings"
        if (Test-Path -LiteralPath $oneDriveSettingsFolderPath) {
            Write-Host "Removing local OneDrive settings folder..." -ForegroundColor Gray
            Remove-Item -Path $oneDriveSettingsFolderPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Orchestrates a full OneDrive reset for the current AVD user session
function Reset-AVDOneDrive {
    [CmdletBinding()]
    param()

    process {
        Write-Verbose "Starting OneDrive reset for $env:USERNAME in Azure Virtual Desktop..."
        Stop-AvdUserOneDriveProcess
        Invoke-AvdOneDriveEngineReset
        Clear-AvdOneDriveUserRegistryCaches
        Remove-AvdOneDriveLocalSettingsFolder
        Write-Host "Success: OneDrive has been reset. Next launch will require sign-in." -ForegroundColor Green
    }
}

Reset-AVDOneDrive -Verbose
