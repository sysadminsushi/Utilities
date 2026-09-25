<#
.SYNOPSIS
    Clears the Office identity cache for the current user.

.DESCRIPTION
    Stops common Office apps, then deletes the Office 16.0 Identity
    registry key so apps prompt for sign-in again.

.NOTES
    Author:  sysadminsushi
    Version: 9.24.2026
#>
function Clear-MicrosoftOfficeIdentityCache {
    Get-Process -Name WINWORD, EXCEL, POWERPNT, OUTLOOK, ONENOTE, MSACCESS -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    $identityKeyPath = "HKCU:\Software\Microsoft\Office\16.0\Common\Identity"

    if (-not (Test-Path $identityKeyPath)) {
        Write-Output "Office identity key not found: $identityKeyPath"
        return
    }

    Remove-Item -Path $identityKeyPath -Recurse -Force
    Write-Output "Cleared Office identity cache: $identityKeyPath"
}

Clear-MicrosoftOfficeIdentityCache
