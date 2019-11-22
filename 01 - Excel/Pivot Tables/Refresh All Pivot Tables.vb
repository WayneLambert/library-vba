'In the worksheet's code window...
Private Sub Worksheet_Change(ByVal Target As Range)
	If Not Application.Intersect(Range(me.ListObjects(1).DataBodyRange,Target)) Is Nothing Then
		Call RefreshAllPivots
	End If
End Sub

'In a standard code module...
Sub RefreshAllPivots()

Dim ws As Worksheet
Dim pt As PivotTable

For Each ws In ThisWorkbook.Worksheets
    For Each pt In ws.PivotTables
        pt.RefreshTable
    Next pt
Next ws

End Sub