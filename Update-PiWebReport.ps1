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

    .PARAMETER OldValue
    The old value that should be replaced in the reports.

    .PARAMETER NewValue
    The new value that will replace the old value.

    .EXAMPLE
    $params = @{
        OriginalReportPath = 'C:\temp\Original'
        ModifiedReportPath = 'C:\temp\Modified'
        TempReportPath = 'C:\temp\Temp'
        OldValue = 'https://pdrs.sdi.corpintra.net'
        NewValue = 'https://pdrs.mo360cp.i.mercedes-benz.com'
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
    try
    {
        foreach ($piwebReport in $piwebReports)
        {
            # Create a temporary folder and copy PiWeb Report
            Copy-Item -Path $piwebReport.FullName -Destination $TempReportPath

            # Expand PiWeb Report
            Expand-PiWebReport -TempReportPath $TempReportPath -Report $piwebReport

            # Search and replace value
            $piwebReportPages = Get-ChildItem -Path "$TempReportPath\PageData"
            foreach ($piwebReportPage in $piwebReportPages)
            {
                Update-ValueInPage -Page $piwebReportPage -OldValue $OldValue -NewValue $NewValue
            }

            # Archive PiWeb Report
            Compress-PiWebReport -TempReportPath $TempReportPath -Report $piwebReport

            # If necessary, create a subfolder and move the PiWeb Report to the modified folder.
            Move-Report -TempReportPath $TempReportPath -ModifiedReportPath $ModifiedReportPath -Report $piwebReport

            # Clean Temp Folder
            Remove-Report -TempReportPath $TempReportPath
        }
    }
    catch
    {
        Write-Host "Program aborted: $_.Exception.Message" -ForegroundColor Red
        Exit 1
    }

    Write-Host 'Replacing values in the PiWeb reports are done.' -ForegroundColor Green
    
    Remove-TempFolder -TempReportPath $TempReportPath
}

$params = @{
    OriginalReportPath = 'C:\temp\original'
    ModifiedReportPath = 'C:\temp\modified'
    TempReportPath = 'C:\temp\temp'
    OldValue = 'https://pdrs.sdi.corpintra.net'
    NewValue = 'https://pdrs.mo360cp.i.mercedes-benz.com'
}

Update-PiWebReport @params