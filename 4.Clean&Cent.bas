Attribute VB_Name = "Module1"
Sub RemoveDuplicatesAndMergeData()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim brand As String
    Dim dict As Object
    Dim currentValue As Variant
    Dim existingValue As Variant
    Dim toDelete As Collection
    Dim currentRow As Long
    Dim cell As Range
    Dim sourceFile As String
    Dim destinationFile As String
    Dim sourceWB As Workbook
    
    ' Prompt to select the source Excel file
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Select the Source Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show = -1 Then
            sourceFile = .SelectedItems(1)
        Else
            MsgBox "No file selected. Exiting."
            Exit Sub
        End If
    End With

    ' Open the selected source workbook
    Set sourceWB = Workbooks.Open(sourceFile)
    Set ws = sourceWB.Sheets(1) ' Assuming the data is in the first sheet
    
    ' Create the dictionary for tracking duplicates
    Set dict = CreateObject("Scripting.Dictionary")
    Set toDelete = New Collection ' For storing rows to delete
    
    ' Find the last row and last column in the used range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Loop through each row starting from the second row (assuming row 1 has headers)
    For i = 2 To lastRow
        brand = ws.Cells(i, 1).Value ' Get the brand name from column A
        
        ' Check if the brand is already in the dictionary
        If Not dict.exists(brand) Then
            ' If brand doesn't exist in dictionary, add it with the row number
            dict.Add brand, i
        Else
            ' If brand exists, process the duplicate row
            currentRow = dict(brand)
            
            ' Loop through columns B to the last column
            For j = 2 To lastCol
                currentValue = ws.Cells(i, j).Value
                existingValue = ws.Cells(currentRow, j).Value
                
                ' Handle numeric columns (e.g., Quantity, Price)
                If IsNumeric(currentValue) And IsNumeric(existingValue) Then
                    ' Keep the larger numeric value
                    If currentValue > existingValue Then
                        ws.Cells(currentRow, j).Value = currentValue
                    End If
                ' Handle text columns (e.g., Color, Brand Name)
                ElseIf Not IsNumeric(currentValue) And Not IsNumeric(existingValue) Then
                    ' Merge text values if they are different (with a comma)
                    If currentValue <> existingValue Then
                        ws.Cells(currentRow, j).Value = existingValue & ", " & currentValue
                    End If
                ' Handle text vs numeric (keep the non-empty value)
                ElseIf Not IsNumeric(currentValue) Then
                    If currentValue <> "" Then
                        ws.Cells(currentRow, j).Value = currentValue
                    End If
                Else
                    If existingValue <> "" Then
                        ws.Cells(currentRow, j).Value = existingValue
                    End If
                End If
            Next j
            
            ' Mark the current duplicate row for deletion
            toDelete.Add i
        End If
    Next i
    
    ' Delete duplicate rows in reverse order (to avoid shifting issues)
    For i = toDelete.Count To 1 Step -1
        currentRow = toDelete(i)
        ws.Rows(currentRow).Delete
    Next i

    ' Prompt to select the destination location and filename
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Save the Merged Data File"
        .FilterIndex = 1
        .InitialFileName = "MergedData.xlsx"
        If .Show = -1 Then
            destinationFile = .SelectedItems(1)
        Else
            MsgBox "No destination file selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Save the processed source workbook to the destination file
    sourceWB.SaveAs destinationFile ' Save the new file to the selected location
    sourceWB.Close SaveChanges:=False ' Close the workbook without saving changes

    MsgBox "Duplicates removed and data merged successfully!", vbInformation
End Sub

