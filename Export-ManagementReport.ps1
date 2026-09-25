<#
.SYNOPSIS
    Generates a timestamped CSV report from an input array.

.DESCRIPTION
    Ensures the target directory exists, builds a unique filename from an
    optional report name and a timestamp, and exports the data to CSV.

.PARAMETER LogDirectory
    Folder where the CSV is saved. Created if it does not exist.

.PARAMETER DataArray
    Objects to export.

.PARAMETER ReportName
    Optional label in the filename (for example SharedMailboxes).

.EXAMPLE
    Export-ManagementReport -LogDirectory "C:\Logs" -DataArray $Results -ReportName "SharedMailboxes"

.NOTES
    Author:  sysadminsushi
    Version: 9.24.2026
#>
function Export-ManagementReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LogDirectory,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$DataArray,

        [string]$ReportName
    )

    if (-not (Test-Path $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }

    $safeName = if ([string]::IsNullOrWhiteSpace($ReportName)) {
        $null
    } else {
        ($ReportName -replace '[\\/:*?"<>|]', '_')
    }

    $uniqueTimeStamp = Get-Date -Format "yyyyMMdd_HHmmssfff"
    $fileName = if ($safeName) {
        "Management_Report_${safeName}_$uniqueTimeStamp.csv"
    } else {
        "Management_Report_$uniqueTimeStamp.csv"
    }

    $reportingPath = Join-Path $LogDirectory $fileName

    if ($null -eq $DataArray -or $DataArray.Count -eq 0) {
        Write-Warning "The provided DataArray was empty. No CSV created."
        return
    }

    try {
        $DataArray | Export-Csv -Path $reportingPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Output $reportingPath
    }
    catch {
        Write-Error "An error occurred during export: $($_.Exception.Message)"
    }
}
