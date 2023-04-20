Option Explicit

Dim objExcel, objWorkbook, objFSO, objOutputFile
Dim strExcelFile, strCSVFile

strExcelFile = "C:\path\to\your\excel_file.xls"
strCSVFile = "C:\path\to\your\csv_file.csv"

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(strExcelFile)

objWorkbook.SaveAs strCSVFile, 6 '6 represents the CSV file format

objWorkbook.Close False 'False means don't save changes
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing


'To run this script via Alteryx, you can use the "Run Command" tool and specify the script file path and arguments.
'
'Here's an example:
'
'Add a "Run Command" tool to your Alteryx workflow.
'
'In the "Command" field, enter the path to the VBS script file, for example: "C:\path\to\your\script.vbs"
'
'In the "Arguments" field, enter the path to the XLS file, for example: "C:\path\to\your\excel_file.xls"
'
'Click on the "Output Options" tab and select "Output File(s)".
'
'In the "Output File Name" field, enter the path to the output CSV file, for example: "C:\path\to\your\csv_file.csv"
'
'Run the workflow and the XLS file will be converted to a CSV file.
