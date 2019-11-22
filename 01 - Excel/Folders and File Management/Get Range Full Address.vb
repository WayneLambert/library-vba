'Gets the complete cell address including the worksheet name. Can also get the workbook's name
Public Function GetFullAddress(ByRef r As Range, Optional bIncludeWorkbookName As Boolean) As String

Dim wb As Workbook
Dim sTmp As String

Set wb = ThisWorkbook

sTmp = Evaluate("ADDRESS(" & r.Row & ", " & r.Column & ",1,1,""" & r.Worksheet.Name & """)")
If (r.Count > 1) Then sTmp = sTmp & ":" & r.Cells(r.Count).Address(RowAbsolute:=True, ColumnAbsolute:=True)
If bIncludeWorkbookName = True Then
    Set wb = ThisWorkbook
    sTmp = Replace(sTmp, "!", "'!", vbTextCompare)
    GetFullAddress = "'[" & wb.Name & "]" & sTmp
Else
    GetFullAddress = sTmp
End If

If Not wb Is Nothing Then Set wb = Nothing

End Function