<#
.SYNOPSIS
Generates configuration files in bulk based on Excel data and a template file.

.DESCRIPTION
This script reads configuration data from an Excel file and uses a template configuration file to generate text-based configuration files. Each row in the Excel file corresponds to a new configuration file.

.NOTES
Dependencies:
- Microsoft Excel (COMObject)

Version: 1.5
Author: Brett Verney
Date: 28 November 2024
#>

# Start timer
$timer = Measure-Command {
    $excel = $null
    $workbook = $null
    try {
        # Constants
        $excelFilePath = "$PWD\data.xlsx"
        $templateFilePath = "$PWD\config_template.cnf"
        $outputDirectory = "$PWD\GeneratedConfigs"

        # Verify template file exists
        if (!(Test-Path -Path $templateFilePath)) {
            throw "Template file not found: $templateFilePath"
        }

        # Read the template configuration file
        $configTemplate = Get-Content $templateFilePath -Raw

        # Ensure the output directory exists
        if (!(Test-Path -Path $outputDirectory)) {
            New-Item -Path $outputDirectory -ItemType Directory | Out-Null
            Write-Host "Created directory: $outputDirectory"
        }

        # Open Excel file in read-only mode
        if (!(Test-Path -Path $excelFilePath)) {
            throw "Excel file not found: $excelFilePath"
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false  # Ensure Excel is hidden
        $workbook = $excel.Workbooks.Open($excelFilePath, [System.Reflection.Missing]::Value, $true)  # Open read-only
        $worksheet = $workbook.Sheets.Item(1)
        $range = $worksheet.UsedRange

        # Validate worksheet data
        if ($range.Rows.Count -lt 2 -or $range.Columns.Count -lt 1) {
            throw "Excel file is missing data or headers."
        }

        # Process each row
        $fileCount = 0
        for ($rowIndex = 2; $rowIndex -le $range.Rows.Count; $rowIndex++) {
            # Get filename from column A
            $filename = $worksheet.Cells.Item($rowIndex, 1).Value2
            if ([string]::IsNullOrWhiteSpace($filename)) {
                Write-Host "Skipping row ${rowIndex}: Blank filename."
                continue
            }

            # Check for blanks in other cells
            $rowDataValid = $true
            for ($colIndex = 2; $colIndex -le $range.Columns.Count; $colIndex++) {
                if ([string]::IsNullOrWhiteSpace($worksheet.Cells.Item($rowIndex, $colIndex).Value2)) {
                    $rowDataValid = $false
                    break
                }
            }

            if (-not $rowDataValid) {
                Write-Host "Skipping row ${rowIndex}: Contains blank cells."
                continue
            }

            # Generate config by replacing placeholders
            $outputConfig = $configTemplate
            for ($colIndex = 1; $colIndex -le $range.Columns.Count; $colIndex++) {
                $placeholder = $worksheet.Cells.Item(1, $colIndex).Value2
                $value = $worksheet.Cells.Item($rowIndex, $colIndex).Value2
                # Handle numeric and null values explicitly
                if ($value -eq $null) {
                    $value = ""  # Replace null values with an empty string
                } elseif ([int]::TryParse($value, [ref]$null)) {
                    $value = $value.ToString()  # Ensure numeric values are treated as strings
                }
                if ($placeholder -and $value) {
                    $outputConfig = $outputConfig.Replace("{$placeholder}", $value)
                }
            }

            # Generate output filename
            $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
            $outputFilename = "$filename-$timestamp.txt"
            $outputFilePath = Join-Path -Path $outputDirectory -ChildPath $outputFilename

            # Save to file
            $outputConfig | Out-File -FilePath $outputFilePath -Encoding UTF8
            Write-Host "Generated file: $outputFilename"

            $fileCount++
        }

        # Close Excel
        $workbook.Close($false)
        $workbook = $null
        $excel.Quit()
        Write-Host "$fileCount configuration files created successfully."

    } catch {
        Write-Error "An error occurred: $_"
    } finally {
        # Ensure Excel is always properly released
        if ($workbook) {
            try {
                $workbook.Close($false) | Out-Null
            } catch {
                Write-Host "Failed to close workbook. Continuing cleanup."
            }
        }
        if ($excel) {
            try {
                $excel.Quit() | Out-Null
            } catch {
                Write-Host "Failed to quit Excel application. Continuing cleanup."
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# Output elapsed time
Write-Host "Execution completed in $($timer.TotalSeconds) seconds."
