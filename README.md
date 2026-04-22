PCXLab Automation Tool

How to use:

1. Open PowerShell
2. Navigate to this folder
3. Run:

   .\main.ps1 -Folder "C:\Your\Input\Folder"

Optional:
   -OutputFolder "C:\Output"

Requirements:
- PowerShell 5.1+
- Microsoft Excel (for .xls conversion)
- Internet (first run to install ImportExcel module)

Output:
- Transformed Excel files will be generated in same folder (or output folder)
- Logs will be stored in /logs folder


FIRST TIME SETUP

Option 1 (Recommended):
Run PowerShell as Administrator and execute:

Set-ExecutionPolicy -Scope CurrentUser RemoteSigned

-------------------------------------

Option 2 (No changes to system):
Run the script using:

powershell.exe -ExecutionPolicy Bypass -File .\main.ps1 -Folder "C:\TEST"