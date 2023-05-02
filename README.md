# WinConfigMaster
PowerShell script used to create bulk text-based configuration files.

## Introduction
This PowerShell script reads an Excel file containing configuration data and generates a text-based configuration file for each row. The output file is created by replacing variables in a template configuration file with the values from the corresponding row in the Excel file. The output file is named after the values specified in column A of the Excel file, with a timestamp appended to it.

## Prerequisites
- PowerShell installed on the computer where you will run the script.
- Microsoft Excel installed on the computer where you will run the script.

## How to use the script
1. Clone or download this repository to your local computer.
2. Open the `config_template.cnf` file and paste in the required configuration. Specify variables within the configuration by surrounding the config in square brackets, i.e. [variable1], [variable2] etc. The file currently includes an example that should be overwritten with the required data.
3. Open the Excel file named `data.xlsx'. The headings in the first row should contain the names of the variables you defined in Step 2. Add the data that you wish to replace the variables under each heading. Again, overwrite the data with your requirements.
4. Run the script by opening PowerShell and navigating to the folder where the script is located. Then, run the following command:
`.\WinConfigMaster.ps1`

This will read the configuration data from data.xlsx and create a text-based configuration file for each row in the Excel file using the template configuration file located at config_template.cnf. The script will ignore rows containing blank cells under a column with a heading specified and not create config files for these rows.

Note: If you see an error message like "File cannot be loaded because the execution of scripts is disabled on this system", you need to enable PowerShell script execution by running the following command in an elevated PowerShell session:

`Set-ExecutionPolicy RemoteSigned`

This will allow scripts to be executed on your computer that are signed by a trusted publisher.

The resulting configuration files will be created in the working directory where the script was executed, with the filename from the value in column A of the Excel file, followed by a timestamp in the format -yyyy-MM-dd-HHmm.

## Limitations
- The script assumes that only one worksheet in the Excel file contains the configuration data.
- The script assumes that the Excel file's first row contains the variables' names that will be replaced in the template file.
- The variables in the template configuration file should be enclosed in square brackets, like [variable].
- The filename in column A of the Excel file should not contain any square brackets.
