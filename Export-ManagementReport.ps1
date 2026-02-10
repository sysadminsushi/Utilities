<#
.SYNOPSIS
    Generates a timestamped CSV report from an input array.

.DESCRIPTION
    A standalone utility for management reporting. It ensures the target directory exists,
    generates a unique filename using a high-precision timestamp, and exports
    the provided data array to a CSV.

.PARAMETER LogDirectory
    The folder where the CSV will be saved. Created automatically if missing.

.PARAMETER DataArray
    The collection of objects to be exported to the CSV.

.EXAMPLE
    Export-ManagementReport -LogDirectory "C:\Logs\Reports" -DataArray $Results

.NOTES
    Author: sysadminsushi
    Version: 2.9.2026
#>
function Export-ManagementReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LogDirectory,
        [Parameter(Mandatory)][array]$DataArray
    )

    # Ensure the directory exists without throwing an error if it already exists
    if (-not (Test-Path $LogDirectory)) { 
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null 
    }

    # Generate Timestamped Filename
    $UniqueTimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $ReportingPath = Join-Path $LogDirectory "Management_Report_$UniqueTimeStamp.csv"

    # Logic Execution & Export
    try {
        if ($null -ne $DataArray -and $DataArray.Count -gt 0) {
            $DataArray | Export-Csv -Path $ReportingPath -NoTypeInformation -Encoding UTF8 -Force
            Write-Host "Success! Management report saved to: $ReportingPath" -ForegroundColor Green
        } else {
            Write-Warning "The provided DataArray was empty. No CSV created."
        }
    }
    catch {
        Write-Error "An error occurred during export: $($_.Exception.Message)"
    }
}