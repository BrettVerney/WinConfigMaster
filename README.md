PowerShell Script for Bulk Configuration File Creation
This PowerShell script reads an Excel file containing configuration data and generates a text-based configuration file for each row in the file. The output file is created by replacing variables in a template configuration file with the values from the corresponding row in the Excel file. The output file is named after the filename in column A of the Excel file, with a timestamp appended to it.

Prerequisites
PowerShell installed on the computer where you will run the script.
Microsoft Excel installed on the computer where you will run the script.
How to use the script
Clone or download this repository to your local computer.

Open the config_template.txt file and replace the placeholder variables enclosed in square brackets with the names of the variables you want to use in your configuration files.

Create an Excel file named data.xlsx in the data folder, containing the configuration data for your files. The first row should contain the names of the variables you defined in step 2.

Run the script by opening PowerShell and navigating to the folder where the script is located. Then, run the following command:

sql
Copy code
.\Create-ConfigFiles.ps1
This will read the configuration data from data.xlsx and create a text-based configuration file for each row in the Excel file, using the template configuration file located at config_template.txt.

Note: If you see an error message like "File cannot be loaded because the execution of scripts is disabled on this system", you need to enable PowerShell script execution by running the following command in an elevated PowerShell session:

javascript
Copy code
Set-ExecutionPolicy RemoteSigned
This will allow scripts to be executed on your computer that are signed by a trusted publisher.

The resulting configuration files will be created in the output folder, with the filename from column A of the Excel file, followed by a timestamp in the format -yyyy-MM-dd-HHmm.

Limitations
The script assumes that there is only one worksheet in the Excel file that contains the configuration data.
The script assumes that the first row in the Excel file contains the names of the variables that will be replaced in the template file.
The variables in the template configuration file should be enclosed in square brackets, like [variable].
The filename in column A of the Excel file should not contain any square brackets.
Author
This script was created by John Doe on May 5, 2023.

License
This project is licensed under the MIT License - see the LICENSE.md file for details.