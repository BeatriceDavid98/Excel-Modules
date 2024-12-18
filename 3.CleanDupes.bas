Attribute VB_Name = "Module1"
Sub ImportWithoutDuplicatesInFirstColumnAndSave()
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim targetWorksheet As Worksheet
    Dim filePath As String
    Dim savePath As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim sourceRange As Range
    Dim cell As Range
    Dim uniqueDict As Object
    Dim targetData As Variant
    Dim targetRow As Long

    ' Create a dictionary object to track unique values
    Set uniqueDict = CreateObject("Scripting.Dictionary")

    ' Step 1: Prompt user to select a file
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select an Excel File")
    
    ' Check if the user clicked Cancel
    If filePath = "False" Then Exit Sub

    ' Step 2: Open the selected workbook
    Set sourceWorkbook = Workbooks.Open(filePath)
    Set sourceWorksheet = sourceWorkbook.Sheets(1)  ' Assumes the data is in the first sheet

    ' Step 3: Identify the last row and last column with data
    lastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, 1).End(xlUp).Row
    lastCol = sourceWorksheet.Cells(1, sourceWorksheet.Columns.Count).End(xlToLeft).Column

    ' Step 4: Create a new workbook to store the cleaned data
    Set targetWorkbook = Workbooks.Add
    Set targetWorksheet = targetWorkbook.Sheets(1) ' Default first sheet in new workbook

    ' Step 5: Copy headers first (row 1)
    sourceWorksheet.Rows(1).Copy Destination:=targetWorksheet.Rows(1)

    ' Step 6: Set up to store unique rows in an array
    targetRow = 2 ' Start from row 2 in the target sheet
    ReDim targetData(1 To lastRow, 1 To lastCol) ' Resize array to hold the data

    ' Step 7: Loop through the rows in column A and check for duplicates
    Dim currentRow As Long
    currentRow = 2 ' Start from row 2 (below the header)

    For Each cell In sourceWorksheet.Range(sourceWorksheet.Cells(2, 1), sourceWorksheet.Cells(lastRow, 1))
        If Not uniqueDict.exists(cell.Value) Then
            ' If the value is not in the dictionary, add it and store the entire row in the array
            uniqueDict.Add cell.Value, Nothing
            For col = 1 To lastCol
                targetData(targetRow, col) = sourceWorksheet.Cells(currentRow, col).Value
            Next col
            targetRow = targetRow + 1 ' Move to the next row in the target sheet
        End If
        currentRow = currentRow + 1
    Next cell

    ' Step 8: Write all unique rows to the target sheet in one go
    targetWorksheet.Range(targetWorksheet.Cells(2, 1), targetWorksheet.Cells(targetRow - 1, lastCol)).Value = targetData

    ' Step 9: Ask user for the location and file name to save the new workbook
    savePath = Application.GetSaveAsFilename(InitialFileName:="CleanedData.xlsx", _
              FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Save Cleaned File")

    ' Check if the user clicked Cancel
    If savePath = "False" Then
        MsgBox "The file was not saved.", vbExclamation
        targetWorkbook.Close False ' Close the new workbook without saving
        Exit Sub
    End If

    ' Step 10: Save the new workbook to the specified location
    targetWorkbook.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    targetWorkbook.Close SaveChanges:=False ' Close the new workbook after saving

    ' Step 11: Close the source workbook without saving
    sourceWorkbook.Close SaveChanges:=False

    ' Step 12: Notify user that the process is complete
    MsgBox "Data imported and saved without duplicates in the first column!", vbInformation
End Sub


