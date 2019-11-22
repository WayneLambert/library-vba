'Uses a For... Next Loop to iterate through the rows in the dataset in reverse order
'The If function determines whether the the row is completely blank
'The .Delete method deletes the blank row. 1 is added to the counter
'Informs user how many rows were deleted
Sub DeleteEmptyRows()

Dim wf As WorksheetFunction
Dim iLastRow As Long, iCounter As Long, i As Long

Application.ScreenUpdating = False

iLastRow = Sheet1.UsedRange.Rows.Count + Sheet1.UsedRange.Rows(1).Row - 1

For i = iLastRow To 1 Step -1
    If wf.WorksheetFunction.CountA(Rows(i)) = 0 Then
        Rows(i).Delete
        iCounter = iCounter + 1
    End If
Next i

MsgBox iCounter & " empty rows were deleted.", vbInformation, "Rows Deleted"

Application.ScreenUpdating = True

End Sub