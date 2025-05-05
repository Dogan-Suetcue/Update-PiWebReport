Function Update-PiWebReport
{
    <#
    .SYNOPSIS
    This function replaces variables on all pages of a PiWeb report.

    .DESCRIPTION
    The script creates a folder named "original" and "modified" in the c:\temp path.
    Copy a copy of the PiWeb reports into the c:\temp\original folder and run the script.
    At the end of execution, you will find the modified reports in the c:\temp\modified path.

    .PARAMETER OriginalReportPath
    The path to the folder containing the original PiWeb reports.

    .PARAMETER ModifiedReportPath
    The path to the folder where the modified PiWeb reports will be moved.

    .PARAMETER TempReportPath
    The path to the temporary folder used for processing the reports.

    .PARAMETER ElementNames
    An array of element names where the values should be replaced.

    .PARAMETER OldValue
    The old value that should be replaced in the reports.

    .PARAMETER NewValue
    The new value that will replace the old value.

    .EXAMPLE
    $params = @{
        OriginalReportPath = 'C:\Temp\Original'
        ModifiedReportPath = 'C:\Temp\Modified'
        TempReportPath = 'C:\Temp\Temp'
        ElementNames = @('Application', 'HRef')
        OldValue = 'https://domain.com'
        NewValue = 'https://new-domain.com'
    }
    Update-PiWebReport @params
    This replaces the URLs in the PiWeb reports according to the specified parameters.

    .NOTES
    Author: Dogan Sütcü
    Created on: 10.04.2025
    Prerequisite: 7-Zip should be installed on the PC.
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $OriginalReportPath,

        [Parameter(Mandatory)]
        [string]
        $ModifiedReportPath,

        [Parameter(Mandatory)]
        [string]
        $TempReportPath,

        [Parameter(Mandatory)]
        [string[]]
        $ElementNames,

        [Parameter(Mandatory)]
        [string]
        $OldValue,

        [Parameter(Mandatory)]
        [string]
        $NewValue
    )
    
    if (-not (Get-Module -Name Utility))
    {
        Import-Module -Name "$PSScriptRoot\Modules\Utility.psm1"
    }

    Test-Prerequisite -OriginalReportPath $OriginalReportPath -ModifiedReportPath $ModifiedReportPath -TempReportPath $TempReportPath

    $piwebReports = Get-ChildItem -Path $OriginalReportPath -Include '*.ptx' -Recurse
    $totalReports = $piwebReports.Count
    $reportIndex = 0
    
    try
    {
        foreach ($piwebReport in $piwebReports)
        {
            # Increase the index for the current report
            $reportIndex++

            # Create a temporary folder and copy PiWeb Report
            Copy-Item -Path $piwebReport.FullName -Destination $TempReportPath

            # Expand PiWeb Report
            Expand-PiWebReport -TempReportPath $TempReportPath -Report $piwebReport

            # Search and replace value
            $piwebReportPages = Get-ChildItem -Path "$TempReportPath\PageData"
            $totalPages = $piwebReportPages.Count
            $pageIndex = 0

            foreach ($piwebReportPage in $piwebReportPages)
            {
                Update-ValueInPage -Page $piwebReportPage -ElementNames $ElementNames -OldValue $OldValue -NewValue $NewValue

                $pageIndex++
                $percentComplete = [Math]::Round(($pageIndex / $totalPages) * 100)
                Write-Progress -Activity "Processing report: $($piwebReport.Name)" -Status "Report $reportIndex of $totalReports, Page $pageIndex of $totalPages" -PercentComplete $percentComplete
                Start-Sleep -Milliseconds 100
            }

            # Archive PiWeb Report
            Compress-PiWebReport -TempReportPath $TempReportPath -Report $piwebReport

            # Move the processed PiWeb Report to the Modified folder.
            Move-ModifiedReport -TempReportPath $TempReportPath -ModifiedReportPath $ModifiedReportPath -Report $piwebReport

            # Delete report files in the Temp folder
            Remove-Report -TempReportPath $TempReportPath
        }
    }
    catch
    {
        Write-Host "Program aborted: $_.Exception.Message" -ForegroundColor Red
        Exit 1
    }
    finally
    {
        # Delete Temp Folder
        Remove-TempFolder -TempReportPath $TempReportPath
    }

    Write-Host 'Replacing values in the PiWeb reports are done.' -ForegroundColor Green
    Write-Host 'Press the Enter key to close the programme.'
    Read-Host
}

$params = @{
    OriginalReportPath = 'C:\Temp\Original'
    ModifiedReportPath = 'C:\Temp\Modified'
    TempReportPath = 'C:\Temp\Temp'
    ElementNames = @('Text', 'HRef', 'Application')
    OldValue = 'Hello_PowerShell'
    NewValue = 'Hello_PiWeb'
}

Update-PiWebReport @params