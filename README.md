## Credit
- "A man who cannot give credit where it is due, cannot be trusted to tell the truth.
    - Uknown
I used [OpenAI's ChatGPT](https://chat.openai.com/) to refactor the original script
- The original script can be found at [n0spaces - Export-TSToExcel](https://github.com/n0spaces/Export-TSToExcel/tree/main)
    - Core functionality of the script originall implemented by @[n0spaces](https://github.com/n0spaces)
    - Copyright (c) 2021 Matt Schwartz 

# Use-TsToExcel.ps1

A PowerShell script to export a Configuration Manager Task Sequence to an Excel sheet for documentation.

## Description

This script exports a Configuration Manager task sequence, obtained either from the `Get-CMTaskSequence` cmdlet (or an exported Task Sequence XML), to an Excel sheet for easy reading and navigation.The script provides various parameters to control the output, such as showing the Excel sheet when complete, including macro buttons for expand/collapse groups, and grouping row outlines without macros.

## Parameters

- `sequencePath`: Path to an exported task sequence XML.
- `exportPath`: Path to save exported Excel file. If not provided, Excel sheet is shown without saving it
- `HideProgress`: Default=`$false` (When set `$true` Script hides the progress bar in the PowerShell window)
- `Show`: Default=`$true` (When set `$false` Script does not show the Excel sheet after it is generated)
- `Macro`: Default=`$true` (When set `$false` Script disables macro buttons to expand/collapse groups in the Excel sheet)
- `Outline`: Default=`$true` (When set `$false` Script disables groups outline in Excel sheet & can be collapsed without macro buttons)

## Usage

1. Launch PowerShell from the admin console and temporarily change the execution policy to allow unsigned scripts:

    ```powershell
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted
    ```

2. Dot source the script:

    ```powershell
    . C:\Path\To\Use-TsToExcel.ps1
    ```
    - Link: [Dot-Sourcing on MSFT Learn](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_scripts?view=powershell-7.3#script-scope-and-dot-sourcing)
    
3. Use the script in combination with the Configuration Manager module:

   This command retrieves the task sequence named "Task Sequence" using the `Get-CMTaskSequence` cmdlet and pipes it to the `Use-TsToExcel` script.
   The script generates an Excel document with the task sequence steps formatted for easy readability.

   - Task Sequence by Name, output Excel to "C:\temp\TS.xlsx", sets all other parameters to default
   
    ```powershell
    Get-CMTaskSequence -Name "Task Sequence" | Use-TsToExcel -exportPath "C:\temp\TS.xlsx"
    ```
   - Task Sequence by Package ID, will not show progress, disables collapsing groups, & prevents showing Excel when done
   
    ```powershell
    Get-CMTaskSequence -PackageID "ABC123" | Use-TsToExcel -exportPath "C:\temp\TS.xlsx" -HideProgress $true -Macro $false -Show $false
    ```
    - Link: [Get-CMTaskSequence on MSFT Learn](https://learn.microsoft.com/en-us/powershell/module/configurationmanager/get-cmtasksequence?view=sccm-ps)
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

### GNU General Public License
This script is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This script is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
