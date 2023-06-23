<#
.SYNOPSIS
    A PowerShell script to export a Configuration Manager task sequence to an Excel sheet for documentation.

.DESCRIPTION
    This script exports a Configuration Manager task sequence to an Excel sheet for easy reading and navigation. 
    The script takes as input the path to an exported task sequence XML and optionally, the path to save the Excel file.
    This script is unsigned, so you may need to temporarily change the execution policy to allow it. 
    You can do this by running the following command in your PowerShell session:
        `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted`
    After changing the execution policy, you can dot source the script:
        `. C:\Path\To\Use-TsToExcel.ps1`
    You can then use the script in combination with the Configuration Manager module:
        `Get-CMTaskSequence -Name "Task Sequence" | Use-TsToExcel -Show`

.PARAMETER sequencePath
    Path to the exported task sequence XML. This parameter is mandatory.

.PARAMETER exportPath
    Path to save the exported Excel file. This parameter is optional. 
    If not provided, the Excel sheet is shown without saving it.

.PARAMETER Show
    If set, the script shows the Excel sheet after it is generated.

.PARAMETER Macro
    If set, the script includes macro buttons to expand/collapse groups in the Excel sheet.

.PARAMETER Outline
    If set, the script groups (outlines) rows in the Excel sheet so they can be expanded/collapsed without the use of macro buttons.

.PARAMETER HideProgress
    If set, the script hides the progress bar in the PowerShell window.

.EXAMPLE
    PS> Get-CMTaskSequence -Name "Task Sequence" | Use-TsToExcel -Show

    This command will first use the `Get-CMTaskSequence` cmdlet from the Configuration Manager module to retrieve 
    the task sequence named "Task Sequence". The task sequence object is then piped to the `Use-TsToExcel` script, 
    which generates an Excel document with the task sequence steps formatted for easy readability. The `-Show` 
    parameter causes the script to display the generated Excel document immediately after it is created. If a 
    path for `-exportPath` is not provided, the Excel document will not be saved.

.EXAMPLE
    PS> Use-TsToExcel -sequencePath "C:\temp\TS.xml" -exportPath "C:\temp\TS.xlsx"

    This command will read the task sequence data from the XML file located at "C:\temp\TS.xml", generate an 
    Excel document with the task sequence steps formatted for easy readability, and save the generated Excel 
    document at "C:\temp\TS.xlsx". The Excel document will not be displayed after it is created.

.CREDITS
    I used OpenAI's ChatGPT to refactor the original script
    - The original script can be found at [n0spaces - Export-TSToExcel](https://github.com/n0spaces/Export-TSToExcel/tree/main).
        - Matt Schwartz @ [n0spaces](https://github.com/n0spaces)
            - Implemented the core functionality of the script
            - Copyright (c) 2021 Matt Schwartz

.NOTES
    Version: 2.0
    Creation Date: 2023-06-22
    Copyright (c) 2023 https://github.com/bentman
    https://github.com/bentman/Use-TsToExcel
#>

# Import necessary libraries
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Bind for standard PowerShell parameter usage
[CmdletBinding()]

# Script parameters
param (
    # Task Sequence object from pipeline
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='TaskSequenceInput')]$TaskSequence,
    # path to exported task sequence XML
    [Parameter(Mandatory=$true, ParameterSetName='PathInput')][string]$sequencePath,
    # path to exported Excel file (optional)
    [Parameter(Mandatory=$false)][string]$exportPath
    [switch]$hideProgress = $false                  # set to $true to hide progress bar
    [switch]$show = $true                           # set to $false to hide Excel and quit after script finishes
    [switch]$macro = $true                          # set to $false to disable Excel macros for collapsing/expanding groups
    [switch]$outline = $true                        # set to $false to disable Excel outline
)

# Define variables
$colorStep = 0xFFFFFF                   # color for step rows (white)
$colorStepDisabled = 0xFFFF00           # color for disabled step rows (yellow)
$colorGroup = 0xC0C0C0                  # color for group rows (gray)
$colorGroupDisabled = 0xFFFF00          # color for disabled group rows (yellow)

# Initialize Excel application
$excel = New-Object -ComObject Excel.Application
$excel.Visible = !$hideProgress
$excel.DisplayAlerts = $false

function ConvertToFriendlyName { # Convert a PascalCase sequence type to a space-delimited string
    param ($Type)
    $Type = $Type.Replace("SMS_TaskSequence_", "").Replace("Action", "")
    switch ($Type) {
        "RunPowerShellScript" { return "Run PowerShell Script" }
        "DisableBitLocker" { return "Disable BitLocker" }
        "EnableBitLocker" { return "Enable BitLocker" }
        "OfflineEnableBitLocker" { return "Pre-provision BitLocker" }
        "AutoApply" { return "Auto Apply Drivers" }
        Default { return [regex]::Replace($Type, "([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))", "`$1 ") }
    }
}

function SetMaxSize { # Helper function to set the maximum size of a row or column
    param (
        $Range,
        $MaxWidth = 0,
        $MaxHeight = 0
    )
    if ($MaxWidth -gt 0) {
        if ($Range.ColumnWidth -gt $MaxWidth) {
            $Range.ColumnWidth = $MaxWidth
        }
    }
    if ($MaxHeight -gt 0) {
        if ($Range.RowHeight -gt $MaxHeight) {
            $Range.RowHeight = $MaxHeight
        }
    }
}

function HandleStep { # Handles TS Step entries
    param (
        $Entry,
        $IndentLevel,
        $Disabled
    )
    $ws.Range("A$CurrentRow").IndentLevel = $IndentLevel
    if ($Disabled) {
        $ws.Range("A$($CurrentRow):F$CurrentRow").Interior.Color = $ColorStepDisabled
        $ws.Range("A$($CurrentRow):F$CurrentRow").Font.Strikethrough = $true
    } else {
        $ws.Range("A$($CurrentRow):F$CurrentRow").Interior.Color = $ColorStep
    }
    $ws.Range("A$($CurrentRow):B$($CurrentRow)").Font.Bold = $false
    $ws.Range("B$CurrentRow").Cells = "Step"
    # get step attributes
    $ID = $Entry.GetAttribute("id")
    $Description = $Entry.GetAttribute("description")
    $Instruction = $Entry.GetAttribute("instruction")
    $Comment = $Entry.GetAttribute("comment")
    # write step data to Excel
    $ws.Range("C$CurrentRow").Cells = $ID
    $ws.Range("D$CurrentRow").Cells = $Description
    $ws.Range("E$CurrentRow").Cells = $Instruction
    $ws.Range("F$CurrentRow").Cells = $Comment
    # increment row
    $CurrentRow++
}

function HandleGroup { # Handles TS Group entries
    param (
        $Entry,
        $IndentLevel,
        $Disabled,
        $ws,
        $Macro,
        $ColorGroupDisabled,
        $ColorGroup,
        $Outline
    )
    # Set indentation based on whether we're in a macro or not
    $indent = if ($Macro) { $IndentLevel + 1 } else { $IndentLevel }
    $ws.Range("A$CurrentRow").IndentLevel = $indent
    # Set color and strikethrough properties based on whether the group is disabled
    if ($Disabled) {
        $ws.Range("A$($CurrentRow):F$CurrentRow").Interior.Color = $ColorGroupDisabled
        $ws.Range("A$($CurrentRow):F$CurrentRow").Font.Strikethrough = $true
    } else {
        $ws.Range("A$($CurrentRow):F$CurrentRow").Interior.Color = $ColorGroup
    }
    $ws.Range("B$CurrentRow").Cells = "Group"
    $ws.Range("A$( $CurrentRow ):B$( $CurrentRow )").Font.Bold = $true
    # Add expand button for macro
    if ($Macro) {
        $top = $ws.Range("A$CurrentRow").Top + 4
        $left = (($IndentLevel - 1) * 7) + 4
        $shape = $ws.Shapes.AddShape(7, $left, $top, 7, 7)
        $shape.Fill.ForeColor.RGB = 0
        $shape.Line.ForeColor.RGB = 0
        $shape.Rotation = 180
        $shape.Name = "ExpandShape$CurrentRow"
    }
    # Recursively call WriteEntry for each child of this group
    $FirstRow = $CurrentRow + 1
    foreach ($Child in $Entry.ChildNodes) {
        if ($Child.LocalName -eq "group" -or $Child.LocalName -eq "step") {
            WriteEntry -Entry $Child -IndentLevel ($IndentLevel + 1) -Disabled $Disabled
        }
    }
    # Group rows for outline
    if ($Outline) {
        $ws.Rows("$($FirstRow):$CurrentRow").Group() | Out-Null
    }
    # Add code for button in macro
    if ($Macro) {
        $SubName = "$($shape.Name)Clicked"
        $Code = "Sub $($SubName)()`n"
        $Code += "ToggleRowsHidden `"$($FirstRow):$CurrentRow`", ActiveSheet.Shapes(`"$($shape.Name)`")`n"
        $Code += "End Sub"
        $Module.CodeModule.AddFromString($Code)
        $shape.OnAction = $SubName
    }
}

function FillEntries { # Populates an Excel worksheet with task sequence steps.
    param (
        $Sequence,
        $IndentLevel,
        $Disabled = $false
    )
    # Loop over each child node in the sequence
    foreach ($Child in $Sequence.ChildNodes) {
        # Check if the child is a group or a step
        if ($Child.LocalName -eq "group" -or $Child.LocalName -eq "step") {
            # If the child is a group, call the HandleGroup function
            if ($Child.LocalName -eq "group") {
                HandleGroup -Entry $Child -IndentLevel $IndentLevel -Disabled $Disabled
            }
            # If the child is a step, call the HandleStep function
            elseif ($Child.LocalName -eq "step") {
                HandleStep -Entry $Child -IndentLevel $IndentLevel -Disabled $Disabled
            }
        }
    }
}

function WriteEntry { # Write a task sequence group or step to the excel sheet
    param (
        $Entry,
        $IndentLevel = 0,
        $Disabled = $false
    )
    if (-not $hideProgress) {
        [int]$progress = (($currentRow - 1)/$totalEntries) * 100
        Write-Progress `
            -Activity "Generating Excel sheet..." `
            -Status "Entry $($currentRow - 1)/$totalEntries ($progress%)" -PercentComplete $progress
    }
    $currentRow++
    # variables
    $ws.Range("A$currentRow").Cells = $Entry.Sequence
    $ws.Range("B$currentRow").Cells = ConvertToFriendlyName -Type $Entry.TypeName
    $ws.Range("C$currentRow").Cells = $Entry.PackageName
    $ws.Range("D$currentRow").Cells = $Entry.PackageID
    $ws.Range("E$currentRow").Cells = $Entry.Description
    # set indent level
    $ws.Range("A$currentRow").IndentLevel = $IndentLevel
    # handle steps
    if ($Entry.LocalName -eq "step") {
        HandleStep -Entry $Entry -IndentLevel $IndentLevel -Disabled $Disabled
    }
    # handle groups
    elseif ($Entry.LocalName -eq "group") {
        HandleGroup -Entry $Entry -IndentLevel $IndentLevel -Disabled $Disabled
    }
}

function ClampSize { # Adjusts the size of a range to not exceed a specified maximum.
    param (
        $Range,
        $MaxWidth = 0,
        $MaxHeight = 0
    )

    if ($MaxWidth -gt 0 -and $Range.ColumnWidth -gt $MaxWidth) {
        $Range.ColumnWidth = $MaxWidth
    }

    if ($MaxHeight -gt 0 -and $Range.RowHeight -gt $MaxHeight) {
        $Range.RowHeight = $MaxHeight
    }
}

function FillEntries { # Fills the Excel sheet with task sequence steps.
    param(
        $Sequence,
        $IndentLevel
    )

    foreach ($Node in $Sequence.ChildNodes) {
        switch ($Node.NodeName) {
            "Group" {
                HandleGroup -Group $Node -IndentLevel $IndentLevel
            }
            "Step" {
                HandleStep -Step $Node -IndentLevel $IndentLevel
            }
        }
    }
}

################# Script body starts here #################

process { 
# If a Task Sequence object is provided, extract the sequence XML and details
if ($TaskSequence) {
    try {
        $Sequence = [System.Xml.XmlDocument]$TaskSequence.Sequence
        $TSName = $TaskSequence.Name
        $PackageID = $TaskSequence.PackageID
    }
    catch {
        Write-Error "`nFailed to extract sequence XML and details from the Task Sequence object."
        Write-Error "`nError: $($_.Exception.Message)"
        return
    }
}
# If no Task Sequence object is provided, load the sequence XML from file
else {
    try {
        [xml]$Sequence = Get-Content -Path $sequencePath
    }
    catch {
        Write-Error "`nFailed to load the sequence XML from file."
        Write-Error "`nError: $($_.Exception.Message)"
        return
    }
}

# Initialize indent level
$IndentLevel = 0

# If the Macro parameter is passed, set indent level to 1
if ($Macro) {
    $IndentLevel = 1
}

    # Call FillEntries function to fill the Excel sheet with task sequence steps
    FillEntries -Sequence $Sequence -IndentLevel $IndentLevel

    # set column sizes
    $ws.Columns("A:F").ColumnWidth = 70
    $ws.Columns.AutoFit()
    $ws.Columns("C").ColumnWidth = 70
    $ws.Columns("E").ColumnWidth = 8.43
    ClampSize -Range $ws.Columns("F") -MaxWidth 100

    for ($i = 3; $i -le $CurrentRow; $i++) {
        ClampSize -Range $ws.Rows("$i") -MaxHeight 40
    }

    # apply gray borders
    $ws.Range("A2:F$CurrentRow").Borders.Color = 0x808080
    $ws.Range("A2:F$CurrentRow").Borders.LineStyle = 1

    # freeze top row
    $ws.Rows("3").Select()
    $excel.ActiveWindow.FreezePanes = $true
    $ws.Range("A1").Select()

    # save and show excel
    if ($ExportPath) {
        $ws.SaveAs($ExportPath.FullName, if ($ExportPath.Extension -eq ".xlsx") { $null } else { 52 })
    }
    $excel.Visible = $Show
    $excel.DisplayAlerts = $true

    # cleanup
    if (-not $excel.Visible) {
        $wb?.Close()
        $excel?.Quit()
    }

    # Loop stored object 
    $ws, $wb, $excel | ForEach-Object {
        if ($_ -ne $null) {
            # Release COM object to free up resources
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($_)  
        }
    }
    # Trigger garbage collection to reclaim memory resources
    [GC]::Collect()  

    # Display progress update indicating completion
    Write-Progress -Activity "Generating Excel sheet..." -Completed 
}
