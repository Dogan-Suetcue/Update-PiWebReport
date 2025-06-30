# Update-PiWebReport

## Overview

The `Update-PiWebReport` PowerShell function is designed to replace specified values within PiWeb reports. It processes each report by copying it to a temporary location, expanding its contents, performing the value replacements, and then archiving the modified report. The final modified reports are stored in a designated folder.

## Prerequisites

- **7-Zip**: Ensure that 7-Zip is installed on the PC. 7-Zip is open-source software licensed under the GNU LGPL. Please review the [7-Zip License](https://www.7-zip.org/license.txt) for more information.
- **Utility Module**: The script imports a custom Utility module from the script's root directory.

## Parameters

- **OriginalReportPath**: The path to the folder containing the original PiWeb reports.
- **ModifiedReportPath**: The path to the folder where the modified PiWeb reports will be moved.
- **TempReportPath**: The path to the temporary folder used for processing the reports.
- **ElementNames**: An array of element names where the values should be replaced.
- **OldValue**: The old value that should be replaced in the reports.
- **NewValue**: The new value that will replace the old value.

## Usage Example

```powershell
$params = @{
    OriginalReportPath = 'C:\Temp\Original'
    ModifiedReportPath = 'C:\Temp\Modified'
    TempReportPath = 'C:\Temp\Temp'
    ElementNames = @('Application', 'HRef')
    OldValue = 'https://domain.com'
    NewValue = 'https://new-domain.com'
}
Update-PiWebReport @params
```

This example replaces the URLs in the PiWeb reports according to the specified parameters.

## Detailed Description

The script performs the following steps:

1. **Create Folders**:

   - It creates folders named "original" and "modified" in the `c:\temp` path.

2. **Copy Reports**:

   - Copies the original PiWeb reports into the `c:\temp\original` folder.

3. **Process Reports**:

   - The script processes each report by:
     - **Copying** it to a temporary folder.
     - **Expanding** the report.
     - **Replacing** specified values in the report pages.
     - **Archiving** the modified report.
     - **Moving** the modified report to the `c:\temp\modified` folder.
     - **Cleaning up** temporary files.

4. **Error Handling**:

   - If an error occurs during processing, the script catches the exception and displays an error message.

5. **Cleanup**:

   - Finally, the script ensures that the temporary folder is deleted after processing is complete.

6. **Completion Message**:
   - Once all reports have been processed, a message is displayed indicating the completion of the value replacements in the PiWeb reports.
