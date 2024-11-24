# Excel SQL Exporter

This tool exports SQL Tables and Views to Excel files and then also optionally uploads them to an FTP site for integrations with cloud systems and provide an easier solution compared with SSIS which will often not work well with large text fields

If you are looking for a tool that works the other way around exporting from Excel into SQL then please see the other project instead:
https://github.com/robinwilson16/ExcelSQLImporter

## Purpose

The tool was created as a replacement for Microsoft SQL Integration Services (SSIS) which can work well with smaller files but these days has a lot of limitations which this tool overcomes:
- Excel columns that contain a large number of characters can be exported without any errors or changes being made to settings
- All rows are evaluated when setting column sizes to avoid errors you get with SSIS when the first rows contain less data than subsequent rows and the column size is set to the maximum size needed (for the importer)
- Data types are detected automatically so will export correctly without code page errors, truncated values, missing values where a column mixes text and numbers
- The tool is simpler to use as just requires .NET 9 runtime to be installed and does not require Excel binaries or data access components or any other special settings

## Prereqisites

You will need to install the Microsoft .NET Runtime 9.0 available from: https://dotnet.microsoft.com/en-us/download/dotnet/9.0
Nothing else needs to be installed as this software can just be unzipped and run.

## Setting Up

Download the latest release from: https://github.com/robinwilson16/ExcelSQLExporter/releases/latest

If you have an Intel/AMD machine (most likely) then pick the `amd64` version but if you have a an ARM Snapdragon device then pick the `arm64` version.

Download and extract the zip file to a folder of your choice.

Now edit the appsettings.json to fill in details for:
| Item | Purpose |
| ---- | ---- |
| Database Connection | Connection to the database to read data |
| Database Table | Where you are getting the data from which should be a table or view |
| Excel File | Where you are saving the data to which should be a file ending in .xlsx |
| FTP Connection (Optional) | Where you are then uploading the saved Excel file to |

Once all settings are entered then just click on `ExcelSQLExporter.exe` to run the program.
If you notice any errors appearing in the window then review these, change the settings file and try again.

## Exporting Multiple Files

By default configuration values are picked up from `appsettings.json` but in case you want to use the tool to export multiple Database Tables or Views then when running from the commandline specify the name of the config file after the .exe so to pick up settings from a config file called `FinanceExport.json` execute:

```
ExcelSQLExporter.exe FinanceExport.json
```

## Setting Up a Schedule

You can just click on may wish to set up a schedule to export one or more Excel files each night and the best way to do this in Windows is to use Task Scheduler which is available in all modern versions of Windows.

Create a new task and name it based on the Excel file it will export so for example:
```
ExcelSQLExporter - Finance Data
```

Pick a user account to run the task against. If you used Windows Authentication in your settings file then you will need to pick a user account with sufficient permissions to read the database table or view you are exporting as well as save the Excel file if it is on a network drive.

On the Triggers tab select a schedule such as each day at 18:00.

On the Actions tab specify the location of the Excel SQL Export tool under Program/script (you can use browse to pick it). It should show as something similar to:
```
D:\ExcelSQLExporter\ExcelSQLExporter.exe
```

Optionally if you are exporting more than one file then enter the name of this into the arguments box - e.g.:
```
UsersTable.json
```
