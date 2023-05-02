<#
.SYNOPSIS
This script creates text-based configuration files in bulk based on an Excel data source.

.DESCRIPTION
This script reads an Excel file containing configuration data and generates a text-based configuration file for each row. The output file is created by replacing variables in a template configuration file with the values from the corresponding row in the Excel file. The output file is named after the values specified in column A of the Excel file, with a timestamp appended to it.

.EXAMPLE
PS C:\> .\WinConfigMaster

This example reads the configuration data from "data.xlsx" and creates a text-based configuration file for each row in the Excel file, using the template configuration file "config_template.txt" located in the same directory.

.NOTES
Dependencies:
- Microsoft Excel (to read the Excel file)

Limitations:
- The script assumes that the first row in the Excel file contains the names of the variables that will be replaced in the template file.
- The script assumes that there is only one worksheet in the Excel file that contains the configuration data.

Assumptions:
- The variables in the template configuration file should be enclosed in square brackets, like [variable].
- The filename in column A of the Excel file should not contain any square brackets.

Version: 1.0	
Author: Brett Verney	
Date: May 5, 2023
#>

# Start timer
$timer = Measure-Command {
    # Read data from config_template.cnf
    $config = Get-Content config_template.cnf | Out-String

    # Load Excel data
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open("$PWD\data.xlsx")
    $worksheet = $workbook.Sheets.Item(1)
    $range = $worksheet.UsedRange

    # Initialize counter
    $fileCount = 0

    # Iterate over each row in the Excel data
    for ($i = 2; $i -le $range.Rows.Count; $i++) {
        # Get the filename from column A
        $filename = $worksheet.Cells.Item($i, 1).Value2
        if (!$filename) {
            continue # Ignore rows with blank filename
        }

        # Check if any cells in the row are blank
        $allCellsValid = $true
        for ($j = 2; $j -le $range.Columns.Count; $j++) {
            $cellValue = $worksheet.Cells.Item($i, $j).Value2
            if (!$cellValue) {
                $allCellsValid = $false
                break
            }
        }

        # If any cells are blank, skip creating a file
        if (!$allCellsValid) {
            Write-Host "Skipping row $i because it contains blank cells."
            continue
        }

        # Create a copy of the config file with variables replaced by Excel data
        $output = $config
        for ($j = 1; $j -le $range.Columns.Count; $j++) {
            $varName = $worksheet.Cells.Item(1, $j).Value2
            $varValue = $worksheet.Cells.Item($i, $j).Value2
            if ($varName -and $varValue) {
                $output = $output.Replace("[$varName]", $varValue)
            }
        }

        # Add timestamp to filename
        $timestamp = Get-Date -Format "-yyyy-MM-dd-HHmm"
        $outputFilename = "$filename$timestamp.txt"

        # Save the output to a file with the updated filename
        $outputPath = "$PWD\$outputFilename"
        $output | Out-File $outputPath

        # Increment counter
        $fileCount++
    }

    # Clean up
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    # Print file count
    Write-Host "$fileCount config files created."
}

# Print time taken
Write-Host "Time taken: $($timer.TotalSeconds) seconds."
