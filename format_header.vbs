Sub FormatExcelHeaders()
    Dim folderPath As String, outputFolderPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim excelApp As Object, workbook As Object, worksheet As Object
    Dim filename As String, outputFilename As String
    
    ' Define the folder paths
    folderPath = "C:\Users\User\Desktop\MBR\1"
    outputFolderPath = "C:\Users\User\Desktop\MBR\format_op"
    
    ' Create the output folder if it does not exist
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputFolderPath) Then
        fso.CreateFolder outputFolderPath
    End If
    
    ' Delete the output folder if it already exists and contains files
    If fso.FolderExists(outputFolderPath) And fso.GetFolder(outputFolderPath).Files.Count > 0 Then
        fso.DeleteFolder outputFolderPath, True
        fso.CreateFolder outputFolderPath
    End If
    
    ' Get the folder object and loop through all files in the folder
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files
        ' Check if the file is an Excel file
        If fso.GetExtensionName(file.Path) Like "xls*" Then
            ' Open the Excel file
            Set excelApp = CreateObject("Excel.Application")
            Set workbook = excelApp.Workbooks.Open(file.Path)
            Set worksheet = workbook.ActiveSheet
            
            ' Get the filename without the extension
            filename = fso.GetBaseName(file.Name)
            outputFilename = outputFolderPath & "\" & filename & ".xlsx"
            
            ' Change the header text direction to vertical
            worksheet.Range("A1:Z1").Orientation = 90
            worksheet.UsedRange.Columns.AutoFit
            
            ' Save and close the output file
            workbook.SaveAs outputFilename, 51
            workbook.Close
            excelApp.Quit
            
            ' Show a message box with the output file path
            MsgBox "Output file saved to " & outputFilename
        End If
    Next file
    
    ' Clean up objects
    Set worksheet = Nothing
    Set workbook = Nothing
    Set excelApp = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Sub

