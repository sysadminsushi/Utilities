<#
.SYNOPSIS
    Clears the Office identity cache for the current user.

.DESCRIPTION
    Removes cached authentication tokens and identity data so that
    Microsoft Office applications prompt for sign‑in again.

.AUTHOR
    sysadminsushi

.VERSION
    2.22.2026
#>

# Clears the Office identity cache by deleting the identity registry key for the current user.
function Clear-MicrosoftOfficeIdentityCache {

    # Delete the Microsoft Office identity registry path for the current user
    reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Identity" /f | Out-Null
}

# Executes the Office identity cache clearing process
Clear-MicrosoftOfficeIdentityCache