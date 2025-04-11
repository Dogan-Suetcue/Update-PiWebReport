Set-Variable -Name "SevenZipPath" -Value 'C:\Program Files\7-Zip\7z.exe' -Option Constant

function Test-7ZipInstallation {
    if (-not (Test-Path -Path $SevenZipPath)) {
        Write-Host "7-Zip is not installed. Please install 7-Zip to continue." -ForegroundColor Red
        Exit 1
    }
}

function New-DirectoryIfNotExists {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $DirectoryPath,

        [Parameter(Mandatory = $false)]
        [string]
        $MessageIfCreated
    )

    if (!(Test-Path -Path $DirectoryPath)) {
        New-Item -Path $DirectoryPath -ItemType Directory
        if ($MessageIfCreated) {
            Write-Host $MessageIfCreated
        }
    }
}

function Test-PiWebReportsExist {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $OriginalReportPath
    )

    $reports = Get-ChildItem -Path $OriginalReportPath -Include '*.ptx' -Recurse
    if ($reports.Length -eq 0) {
        Write-Host "To continue the program, please save the copy of the original PiWeb reports in the directory '$($OriginalReportPath)' and run the script again"
        Exit 1
    }
}

function Test-Prerequisite {
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
        $TempReportPath
    )

    Test-7ZipInstallation

    $params = @{
        DirectoryPath = $OriginalReportPath
        MessageIfCreated = "Created missing directory '$($OriginalReportPath)'. Please save the copy of the original PiWeb reports in this folder and run the script again."
    }
    New-DirectoryIfNotExists @params

    Test-PiWebReportsExist -OriginalReportPath $OriginalReportPath

    New-DirectoryIfNotExists -DirectoryPath $ModifiedReportPath

    New-DirectoryIfNotExists -DirectoryPath $TempReportPath
}

function Expand-PiWebReport
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [string]
        $TempReportPath,

        [Parameter(Mandatory)]
        [System.IO.FileSystemInfo]
        $Report
    )

    $argumentList = "x `"$(Join-Path -Path $TempReportPath -ChildPath $Report.Name)`" -o$TempReportPath"
    Start-Process -FilePath $SevenZipPath -ArgumentList $argumentList -Wait
}

function Compress-PiWebReport
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $TempReportPath,

        [Parameter(Mandatory)]
        [System.IO.FileSystemInfo]
        $Report
    )

    $argumentList = "u `"$(Join-Path -Path $TempReportPath -ChildPath $Report.Name)`" $TempReportPath\PageData"
    Start-Process -FilePath $SevenZipPath -ArgumentList $argumentList -Wait
}

function Move-Report
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $TempReportPath,

        [Parameter()]
        [string]
        $ModifiedReportPath,

        [Parameter()]
        [System.IO.FileSystemInfo]
        $Report
    )

    $sourcePath = $(Join-Path -Path $TempReportPath -ChildPath $Report.Name)
    $destination = $ModifiedReportPath
    if (!($Report.Directory.Name -eq 'Original'))
    {      
        $path = Join-Path -Path $ModifiedReportPath -ChildPath $Report.Directory.Name
        New-Item -Path $path -ItemType Directory
        $destination = $path
    }                             
    Move-Item -Path $sourcePath -Destination $destination
}

function Update-ValueInPage
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.IO.FileSystemInfo]
        $Page,

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

    try 
    {
        [xml]$content = Get-Content -Path $Page.FullName
        foreach ($elementName in $ElementNames)
        {
            $elements = $content.SelectNodes("//$($elementName)")
            foreach ($element in $elements)
            {
                $currentInnerText = $element.InnerText
                $newInnerText = $currentInnerText -replace $OldValue, $NewValue
                $element.InnerText = $newInnerText
            }
        }

        $content.Save($page.FullName)
    }
    catch 
    {
        throw
    }
}

function Remove-Report
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $TempReportPath
    )
    
    Get-ChildItem -Path $TempReportPath -Include * -File -Recurse | ForEach-Object { $_.Delete() }
}

function Remove-TempFolder
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $TempReportPath
    )

    Remove-Item -Path $TempReportPath -Recurse -Force
}

Export-ModuleMember -Function Test-Prerequisite, Expand-PiWebReport, Compress-PiWebReport, Move-Report, Remove-Report, Update-ValueInPage, Remove-TempFolder
