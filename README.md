## "Who cannot give credit where due, cannot be trusted." - Uknown
- Core functionality of this script was originally implemented by [Matt Schwartz (@n0spaces)](https://github.com/n0spaces)
    - The original script can be found at [git/n0spaces/Export-TSToExcel](https://github.com/n0spaces/Export-TSToExcel/tree/main)
- I used [OpenAI's ChatGPT](https://chat.openai.com/) to assist refactoring the original script

# Use-TsToExcel.ps1

A PowerShell script to export a Configuration Manager Task Sequence to an Excel sheet for documentation.

## Description

This script exports a Configuration Manager task sequence, obtained either from the `Get-CMTaskSequence` cmdlet (or an exported Task Sequence XML), to an Excel sheet for easy reading and navigation.The script provides various parameters to control the output, such as showing the Excel sheet when complete, including macro buttons for expand/collapse groups, and grouping row outlines without macros.

## Parameters

- `exportPath`: Path to save exported Excel file. If not provided, Excel sheet is shown without saving it
- `sequencePath`: Path to an exported task sequence XML when not using `Get-CMTaskSequence` from pipe
- `HideProgress`: Default `$false`... When set `$true` Script hides the progress bar in the PowerShell window
- `Show`: Default `$true`... When set `$false` Script does not show the Excel sheet after it is generated
- `Macro`: Default `$true`... When set `$false` Script disables macro buttons to expand/collapse groups in the Excel sheet
- `Outline`: Default `$true`... When set `$false` Script disables groups outline in Excel sheet & can be collapsed without macro buttons

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

   This command retrieves the task sequence by name with `Get-CMTaskSequence` and pipes it to `Use-TsToExcel`.
   
   - Get TS by Name, output Excel to "C:\temp\TS.xlsx", sets all other parameters to default
   
    ```powershell
    Get-CMTaskSequence -Name "Task Sequence" | Use-TsToExcel -exportPath "C:\temp\TS.xlsx"
    ```
   - Get TS by Pkg ID, doesn't show progress, disables collapse group 'macro', & prevents 'show' Excel when done
   
    ```powershell
    Get-CMTaskSequence -PackageID "ABC123" | Use-TsToExcel -exportPath "C:\temp\TS.xlsx" -HideProgress $true -Macro $false -Show $false
    ```
    - Link: [Get-CMTaskSequence on MSFT Learn](https://learn.microsoft.com/en-us/powershell/module/configurationmanager/get-cmtasksequence?view=sccm-ps)
5. Alternatively, you can use the script with an exported task sequence XML:

   This command reads the task sequence data from exported TS.xml file located at "C:\temp\TS.xml", generates an Excel sheet with  task sequence groups/steps, saves generated Excel sheet to "C:\temp\TS.xlsx", & Excel sheet will not be displayed on completion.
   
    ```powershell
    Use-TsToExcel -sequencePath "C:\temp\TS.xml" -exportPath "C:\temp\TS.xlsx"
    ```
   NOTE: The '-sequencePath' parameter is mandatory when using an exported TS.xml instead of piping results from Get-CMTaskSequence.

## Requirements

- PowerShell (Version 5.1+)
- Microsoft Excel (version 2019+)
- Microsoft Configuration Manager Console
- Microsoft Configuration Manager Module

### Contributions

Contributions are welcome! Please open an issue or submit a pull request if you have suggestions or enhancements.

### License

This script is distributed without any warranty; use at your own risk.
This project is licensed under the GNU General Public License v3. 
See [GNU GPL v3](https://www.gnu.org/licenses/gpl-3.0.html) for details.

