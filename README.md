# Use-TsToExcel.ps1

A PowerShell script to export a Configuration Manager task sequence to an Excel sheet for documentation.

## Credit Original Author

- I used OpenAI ChatGPT to refactor the original script by [n0spaces - Export-TSToExcel](https://github.com/n0spaces/Export-TSToExcel/tree/main).

## Description

This script exports a Configuration Manager task sequence to an Excel sheet for easy reading and navigation. It takes as input the path to an exported task sequence XML and optionally, the path to save the Excel file. The script provides various parameters to control the output, such as showing the Excel sheet, including macro buttons for expand/collapse groups, and grouping rows without macros.

## Requirements

- PowerShell (tested on version 5.1)
- Microsoft Excel (tested on version 2019)
- Microsoft Endpoint Configuration Manager Console

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

    ```powershell
    Get-CMTaskSequence -Name "Task Sequence" | Use-TsToExcel -Show
    ```

    This command retrieves the task sequence named "Task Sequence" using the `Get-CMTaskSequence` cmdlet and pipes it to the `Use-TsToExcel` script. The script generates an Excel document with the task sequence steps formatted for easy readability. The `-Show` parameter causes the script to display the generated Excel document immediately after it is created.

4. Alternatively, you can use the script with task sequence XML:

    ```powershell
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

## Contributions

Contributions are welcome. Please open an issue or submit a pull request.

### GNU General Public License
This script is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This script is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
