<#
.SYNOPSIS
    Performs a full hard reset of Microsoft Outlook (Classic and New).

.DESCRIPTION
    Closes all Outlook processes, removes cached data from multiple locations,
    deletes and recreates the Outlook profile registry key, and restores Outlook
    to a factory‑state launch condition.

.AUTHOR
    sysadminsushi

.VERSION
    2.22.2026
#>

# Performs a complete hard reset of Microsoft Outlook by clearing cached data and rebuilding the profile registry key.
function Outlook-HardReset {

    # Terminate all Microsoft Outlook processes, including New Outlook (Olk)
    Stop-Process -Name "Outlook", "olk" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1

    # Remove Local AppData Outlook cache
    try {
        Remove-Item -Path "$env:LocalAppData\Microsoft\Outlook\*" -Recurse -Force -ErrorAction SilentlyContinue
    } catch {}

    # Remove Roaming AppData Outlook settings and signatures
    try {
        Remove-Item -Path "$env:AppData\Microsoft\Outlook\*" -Recurse -Force -ErrorAction SilentlyContinue
    } catch {}

    # Remove New Outlook (Olk) data
    try {
        Remove-Item -Path "$env:LocalAppData\Microsoft\Olk\*" -Recurse -Force -ErrorAction SilentlyContinue
    } catch {}

    # Delete and recreate the Outlook profile registry key
    $outlookProfileRegistryPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook"

    if (Test-Path $outlookProfileRegistryPath) {
        Remove-Item -Path $outlookProfileRegistryPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    New-Item -Path $outlookProfileRegistryPath -Force | Out-Null
}

# Executes the Outlook hard reset process
Outlook-HardReset
