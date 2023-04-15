Sub FormatExcelHeaders_new()
    Dim inputFolderPath As String, outputFolderPath As String
    Dim fso As Object, inputFolder As Object, inputFile As Object
    Dim excelApp As Object, excelWorkbook As Object, excelWorksheet As Object
    Dim filename As String, outputFilename As String
    
    ' Define the folder paths
    inputFolderPath = "C:\Users\User\Desktop\MBR"
    outputFolderPath = "C:\Users\User\Desktop\MBR\format_op"
    
    ' Delete the output folder if it already exists
    If Dir(outputFolderPath, vbDirectory) <> "" Then
        Kill outputFolderPath & "\*.*"
        RmDir outputFolderPath
    End If
    
    ' Create the output folder
    MkDir outputFolderPath
    
    ' Create a FileSystemObject and get the folder object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set inputFolder = fso.GetFolder(inputFolderPath)
    
    ' Loop through all files in the folder
    For Each inputFile In inputFolder.Files
        ' Check if the file is an Excel file
        If LCase(fso.GetExtensionName(inputFile.Path)) Like "xls*" Then
            ' Open the Excel file
            Set excelApp = CreateObject("Excel.Application")
            Set excelWorkbook = excelApp.Workbooks.Open(inputFile.Path)
            Set excelWorksheet = excelWorkbook.ActiveSheet
            
            ' Get the filename without the extension
            filename = fso.GetBaseName(inputFile.Name)
            outputFilename = outputFolderPath & "\" & filename & ".xlsx"
            
            ' Change the header text direction to vertical
            With excelWorksheet.Range("A1:Z1")
                .Orientation = 90
                .WrapText = False
                .HorizontalAlignment = -4108 ' xlCenter
                .VerticalAlignment = -4108 ' xlCenter
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = 2 ' xlContext
                .MergeCells = False
            End With
            
            ' Autofit the used cells in the worksheet
            excelWorksheet.UsedRange.Columns.AutoFit
            
            ' Save and close the output file
            excelWorkbook.SaveAs outputFilename, 51 ' xlOpenXMLWorkbook
            excelWorkbook.Close
            excelApp.Quit
            
            ' Show a message box with the output file path
            MsgBox "Output file saved to " & outputFilename
        End If
    Next inputFile
    
    ' Clean up objects
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Set inputFolder = Nothing
    Set fso = Nothing
End Sub

