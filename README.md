I apologize for the formatting issue. Here's the entire content as a single script block that you can copy and paste into your README.md file:

```markdown
<#
.SYNOPSIS
    Use-TsToExcel.ps1

.DESCRIPTION
    A PowerShell script to export a Configuration Manager task sequence to an Excel sheet for documentation.

.NOTES
    This script is a refactored version of the original script by [n0spaces](https://github.com/n0spaces/Export-TSToExcel/tree/main).

## Use-TsToExcel.ps1

A PowerShell script to export a Configuration Manager task sequence to an Excel sheet for documentation.

## Description

This script exports a Configuration Manager task sequence to an Excel sheet for easy reading and navigation. The script takes as input the path to an exported task sequence XML and optionally, the path to save the Excel file. The script provides various parameters to control the output, such as showing the Excel sheet, including macro buttons for expand/collapse groups, and grouping rows without macros.

The script requires PowerShell (tested on version 5.1), Microsoft Excel (tested on version 2019), and the Microsoft Endpoint Configuration Manager Console. Make sure you have the necessary permissions to run unsigned scripts on your system.

## Usage

First, launch PowerShell from the admin console and temporarily change the execution policy to allow unsigned scripts:

```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted
```

Dot source the script:

```
. C:\Path\To\Use-TsToExcel.ps1
```

### Using the Configuration Manager module

To generate an Excel sheet from a task sequence using the Configuration Manager module, you can use the following command:

```
Get-CMTaskSequence -Name "Task Sequence" | Use-TsToExcel -Show
```

This command retrieves the task sequence named "Task Sequence" using the `Get-CMTaskSequence` cmdlet and pipes it to the `Use-TsToExcel` script. The script generates an Excel document with the task sequence steps formatted for easy readability. The `-Show` parameter causes the script to display the generated Excel document immediately after it is created.

### Using task sequence XML

To generate an Excel sheet from a task sequence XML file, you can use the following command:

```
Use-TsToExcel -sequencePath "C:\temp\TS.xml" -exportPath "C:\temp\TS.xlsx"
```

This command reads the task sequence data from the XML file located at "C:\temp\TS.xml", generates an Excel document with the task sequence steps formatted for easy readability, and saves the generated Excel document at "C:\temp\TS.xlsx". The Excel document will not be displayed after it is created.

## Parameters

- `sequencePath`: Path to the exported task sequence XML. This parameter is mandatory.
- `exportPath`: Path to save the exported Excel file. This parameter is optional. If not provided, the Excel sheet is shown without saving it.
- `Show`: If set, the script shows the Excel sheet after it is generated.
- `Macro`: If set, the script includes macro buttons to expand/collapse groups in the Excel sheet.
- `Outline`: If set, the script groups (outlines) rows in the Excel sheet so they can be expanded/collapsed without the use of macro buttons.
- `HideProgress`: If set, the script hides the progress bar in the PowerShell window.

## Notes

This script is a refactored version of the original script by [n0spaces](https://github.com/n0spaces/Export-TSToExcel/tree/main).
#>
```

Please make sure to double-check the formatting after copying it to ensure it appears correctly in your GitHub README.md
