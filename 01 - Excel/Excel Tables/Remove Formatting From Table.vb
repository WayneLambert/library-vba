'Removes any formatting from a table
Sub RemoveFormattingFromTable()

Dim MyNormal As Style
Dim sTabName As String, sTabStyle As String

On Error Resume Next
sTabName = ActiveCell.ListObject.Name
If sTabName = vbNullString Then Exit Sub
sTabStyle = ActiveSheet.ListObjects(sTabName).TableStyle
On Error GoTo 0

ActiveWorkbook.Styles.Add "MyNormal"
Set MyNormal = ActiveWorkbook.Styles("MyNormal")
MyNormal.IncludeNumber = False
    
With ActiveSheet.ListObjects(sTabName)
    .Range.Style = "MyNormal"
    .TableStyle = sTabStyle
End With

ActiveWorkbook.Styles("MyNormal").Delete
    
End Sub