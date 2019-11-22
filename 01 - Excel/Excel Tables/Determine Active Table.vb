Sub DetermineActiveTable()

Dim sTbl As String
Dim loActiveTable As ListObject
Dim bInTable As Boolean

bInTable = Not ActiveCell.ListObject Is Nothing

'Determines if the ActiveCell is inside the table
On Error GoTo NoTableSelected
    sTbl = ActiveCell.ListObject.Name
    Set loActiveTable = ActiveSheet.ListObject(sTbl)
On Error GoTo 0

'ActiveSheet.ListObjects(sTbl).ListColumns(ActiveCell.Column).DataBodyRange.Select
loActiveTable.ListColumns(ActiveCell.Column).DataBodyRange.Select

NoTableSelected:
    MsgBox "There is no Excel Table currently selected!", vbCritical, "No Table Found"
    Exit Sub
    
End Sub