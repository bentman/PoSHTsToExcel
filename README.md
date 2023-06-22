# Use-TsToExcel.ps1

A PowerShell script to export a Configuration Manager task sequence to an Excel sheet for documentation.

## Description

This script exports a Configuration Manager task sequence to an Excel sheet for easy reading and navigation. It takes as input the path to an exported task sequence XML and optionally, the path to save the Excel file. The script provides various parameters to control the output, such as showing the Excel sheet, including macro buttons for expand/collapse groups, and grouping rows without macros.

## Parameters

- `sequencePath`: Path to an exported task sequence XML.
- `exportPath`: This parameter is recommended. Path to save exported Excel file. If not provided, Excel sheet is shown without saving it.
- `HideProgress`: Default=`$false` | `$true`= Script hides the progress bar in the PowerShell window.
- `Show`: Default=`$true` |`$false`= Script does not show the Excel sheet after it is generated.
- `Macro`: Default=`$true`|`$false`= Script disables macro buttons to expand/collapse groups in the Excel sheet.
- `Outline`: Default=`$true`|`$false`= Script disables groups outline in Excel sheet so can be collapsed without macro buttons.

## Usage

1. Launch PowerShell from the admin console and temporarily change the execution policy to allow unsigned scripts:

    ```powershell
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted
    ```

2. Dot source the script:

    ```powershell
    . C:\Path\To\Use-TsToExcel.ps1
    ```

3. Use the script in combination with the Configuration Manager module:

   This command retrieves the task sequence named "Task Sequence" using the `Get-CMTaskSequence` cmdlet and pipes it to the `Use-TsToExcel` script.
   The script generates an Excel document with the task sequence steps formatted for easy readability.
   
    ```powershell
    Get-CMTaskSequence -Name "Task Sequence" | Use-TsToExcel -exportPath "C:\temp\TS.xlsx"
    ```
    Task Sequence by Name, output Excel to "C:\temp\TS.xlsx", sets all other parameters to default
   
    ```powershell
    Get-CMTaskSequence -PackageID "ABC123" | Use-TsToExcel -exportPath "C:\temp\TS.xlsx" -HideProgress $true -Macro $false -Show $false
    ```
    Task Sequence by Package ID, will not show progress, disables collapsing groups, & prevents showing Excel when done

5. Alternatively, you can use the script with an exported task sequence XML:

   This command reads the task sequence data from the XML file located at "C:\temp\TS.xml", generates an Excel document with the task sequence steps formatted for easy readability, and saves the generated Excel document at "C:\temp\TS.xlsx". The Excel document will not be displayed after it is created.
   
    ```powershell
    Use-TsToExcel -sequencePath "C:\temp\TS.xml" -exportPath "C:\temp\TS.xlsx"
    ```

   NOTE: The '-sequencePath' parameter is mandatory when using an exported TS.xml instead of piping results from Get-CMTaskSequence.

## Requirements

- PowerShell (Version 5.1+)
- Microsoft Excel (version 2019+)
- Microsoft Configuration Manager Console
- Microsoft Configuration Manager Module

## Contributions

Contributions are welcome. Please open an issue or submit a pull request.

## Credits

- Credit for concept and original script goes to [n0spaces] https://github.com/n0spaces
- The original script can be found at [n0spaces - Export-TSToExcel](https://github.com/n0spaces/Export-TSToExcel/tree/main).
- I used OpenAI's ChatGPT to refactor the original script

### GNU General Public License
This script is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This script is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
