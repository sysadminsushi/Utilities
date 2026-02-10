<#
.SYNOPSIS
    Generates a timestamped CSV report from an input array.

.DESCRIPTION
    A standalone utility for management reporting. It ensures the target directory exists,
    generates a unique filename using an optional report name and a high-precision 
    timestamp, and exports the provided data array to a CSV.

.PARAMETER LogDirectory
    The folder where the CSV will be saved. Created automatically if missing.

.PARAMETER DataArray
    The collection of objects to be exported to the CSV.

.PARAMETER ReportName
    An optional string to identify the report type (e.g., "SharedMailboxes", "UserAudit").
    This is inserted into the filename.

.EXAMPLE
    Export-ManagementReport -LogDirectory "C:\Logs" -DataArray $Results -ReportName "SharedMailboxes"

.NOTES
    Author: sysadminsushi
    Version: 2.9.2026
#>
function Export-ManagementReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LogDirectory,
        [Parameter(Mandatory)][array]$DataArray,
        [Parameter(Mandatory=$false)][string]$ReportName
    )

    # Ensure the directory exists
    if (-not (Test-Path $LogDirectory)) { 
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null 
    }

    # Generate Timestamped Filename
    $UniqueTimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # Logic: If ReportName is provided, include it; otherwise, just use the timestamp.
    if (-not [string]::IsNullOrWhiteSpace($ReportName)) {
        $FileName = "Management_Report_$($ReportName)_$UniqueTimeStamp.csv"
    } else {
        $FileName = "Management_Report_$UniqueTimeStamp.csv"
    }

    $ReportingPath = Join-Path $LogDirectory $FileName

    # Logic Execution & Export to csv
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