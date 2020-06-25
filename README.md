# Excel-Access Database Tool
Tool for transferring tables between Microsoft Access and Excel and clearing Access tables. Performs Access operations many times faster than is possible when done manually.

WHAT IT DOES:

Allows the efficient exporting of tables between Microsoft Access and Excel. 

HOW TO USE:

To test, download the excelAccerssDbTool.xlsm, as the entire macro is embedded inside of it.

Open your VBA editor from the file go to Tools -> References and add the following libraries:
Microsoft Access 16.0 Object Library
Microsoft ActiveX Data Objects 6.1 Library
Microsoft DAO 3.6 Object Library

To start application, run the runUserForm sub. The database tool will open. If you don't have an Access databasee open, the application will prompt you to open one. The name of the database that you've accessed will be displayed in the centree of the user form.

The first "Clear Access Table" tab allows you to clear all rows from any table in the database. To clear a table, select the name of the table from the dropdown list and click "Clear Access Table".

The second "Excel --> Access" tab allows you to transfer whichever Excel worksheet you currently have open into an Access database. You must indicate whether the table has headers. If it does, they must be identical to those in the Access table.

The third "Access --> Excel" tab allows you to transfer an Access table to whichever excel worksheet you currently have open.

SOME NOTES ON RUNNING:

When importing a table from Excel to access, that Excel book must be saved down.
