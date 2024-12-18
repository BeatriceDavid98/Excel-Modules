Attribute VB_Name = "Module1"
Sub SortFilesByColumns()
    Dim folderPath As String
    Dim mainFolderPath As String
    Dim fileDialog As fileDialog
    Dim fso As Object
    Dim file As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim targetFolder As String
    Dim tempFileName As String
    Dim col As Long

    ' Initialize FileSystemObject for working with folders and files
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create a FileDialog to let the user choose the folder with Excel files
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    fileDialog.Title = "Select the Folder with Excel Files"
    
    ' Show the folder picker to select the folder containing the Excel files
    If fileDialog.Show = -1 Then
        folderPath = fileDialog.SelectedItems(1) & "\"
    Else
        MsgBox "No folder selected. Exiting."
        Exit Sub
    End If
    
    ' Create a FileDialog to let the user choose the main folder where subfolders will be created
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    fileDialog.Title = "Select the Main Folder to Create Subfolders"
    
    ' Show the folder picker to select the main folder
    If fileDialog.Show = -1 Then
        mainFolderPath = fileDialog.SelectedItems(1) & "\"
    Else
        MsgBox "No main folder selected. Exiting."
        Exit Sub
    End If
    
    ' Check if the main folder exists
    If Not fso.FolderExists(mainFolderPath) Then
        MsgBox "The specified main folder does not exist."
        Exit Sub
    End If
    
    ' Loop through each file in the folder containing the Excel files
    For Each file In fso.GetFolder(folderPath).Files
        ' Only process Excel files (xlsx extension)
        If LCase(Right(file.Name, 4)) = "xlsx" Then
            ' Try to open the workbook in read-only mode
            Set wb = Workbooks.Open(file.Path, ReadOnly:=True)
            
            ' Check if the workbook opened successfully
            If Not wb Is Nothing Then
                ' Access the first worksheet
                Set ws = wb.Sheets(1)
                
                ' Find the last row with data in the worksheet
                lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

                ' Find the last column with data in the first row up to the last row
                lastColumn = 1
                For col = 1 To ws.Columns.Count
                    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col))) > 0 Then
                        lastColumn = col
                    Else
                        Exit For
                    End If
                Next col

                ' Create a subfolder named after the column count inside the main folder
                targetFolder = mainFolderPath & lastColumn & "\"
                If Not fso.FolderExists(targetFolder) Then
                    fso.CreateFolder targetFolder
                End If

                ' Copy the file to the folder based on column count
                tempFileName = fso.GetFileName(file.Path)
                fso.CopyFile file.Path, targetFolder & tempFileName

                ' Close the workbook without saving
                wb.Close False
                Set wb = Nothing
            Else
                MsgBox "Failed to open file: " & file.Name
            End If
        End If
    Next file

    MsgBox "Files have been sorted and copied successfully!"
End Sub

