'Returns TRUE if sheet exists in the active workbook
Private Function SheetExists(sName) As Boolean

Dim x As Object
On Error Resume Next
Set x = ActiveWorkbook.Sheets(sName)

If Err = 0 Then SheetExists = True Else SheetExists = False

End Function