Attribute VB_Name = "Module2"
Sub CombineExcelFiles()

    Dim FolderPath As String
    Dim Filename As String
    Dim WorkbookSource As Workbook
    Dim WorksheetSource As Worksheet
    Dim WorkbookDest As Workbook
    Dim WorksheetDest As Worksheet
    Dim LastRowDest As Long
    Dim LastRowSource As Long
    Dim LastColSource As Long
    Dim SaveFileDialog As fileDialog
    Dim SavePath As String
    Dim FolderDialog As fileDialog

    ' Prompt user to select the folder containing Excel files to combine
    Set FolderDialog = Application.fileDialog(msoFileDialogFolderPicker)
    FolderDialog.Title = "Select Folder Containing Excel Files"
    If FolderDialog.Show = -1 Then
        FolderPath = FolderDialog.SelectedItems(1) & "\"
    Else
        MsgBox "No folder selected. Exiting..."
        Exit Sub
    End If

    ' Create a new workbook for the combined data
    Set WorkbookDest = Workbooks.Add
    Set WorksheetDest = WorkbookDest.Sheets(1)

    ' Prompt user to select the location and filename for the new combined file
    Set SaveFileDialog = Application.fileDialog(msoFileDialogSaveAs)
    SaveFileDialog.Title = "Save Combined File As"
    SaveFileDialog.FilterIndex = 1
    If SaveFileDialog.Show = -1 Then
        SavePath = SaveFileDialog.SelectedItems(1)
    Else
        MsgBox "No location selected. Exiting..."
        Exit Sub
    End If

    ' Ensure the user has selected a valid file type (xlsx)
    If Right(SavePath, 5) <> ".xlsx" Then
        SavePath = SavePath & ".xlsx"
    End If

    ' Get the first file in the folder
    Filename = Dir(FolderPath & "*.xlsx")

    ' Loop through each file in the folder
    Do While Filename <> ""
        ' Open the source workbook
        Set WorkbookSource = Workbooks.Open(FolderPath & Filename)
        Set WorksheetSource = WorkbookSource.Sheets(1) ' Assuming data is on the first sheet

        ' Find the last row and last column in the source worksheet
        LastRowSource = WorksheetSource.Cells(Rows.Count, 1).End(xlUp).row
        LastColSource = WorksheetSource.Cells(1, Columns.Count).End(xlToLeft).Column

        ' Find the last row in the destination worksheet
        If WorksheetDest.Cells(1, 1).Value = "" Then
            ' If destination is empty, copy the headers from the first file
            WorksheetSource.Rows(1).Copy Destination:=WorksheetDest.Rows(1)
            LastRowDest = 1
        Else
            ' Otherwise, find the last used row in the destination
            LastRowDest = WorksheetDest.Cells(Rows.Count, 1).End(xlUp).row
        End If

        ' Copy the data (excluding headers) from the source to the destination
        WorksheetSource.Range(WorksheetSource.Cells(2, 1), WorksheetSource.Cells(LastRowSource, LastColSource)).Copy _
            Destination:=WorksheetDest.Cells(LastRowDest + 1, 1)

        ' Close the source workbook without saving
        WorkbookSource.Close False

        ' Get the next file
        Filename = Dir
    Loop

    ' Auto-fit columns in the destination worksheet
    WorksheetDest.Columns.AutoFit

    ' Save the combined file to the user-specified path
    WorkbookDest.SaveAs SavePath

    ' Close the destination workbook
    WorkbookDest.Close

    ' Notify user
    MsgBox "Files have been successfully combined and saved as " & SavePath, vbInformation

End Sub


