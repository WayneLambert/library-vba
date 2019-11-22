'Determines whether either an Excel table or a range is currently filtered
'irrespective of whether the range belongs to an Excel table
Function IsDataFiltered(ByRef r As Range) As Boolean

Dim iRC As Long, iFC As Long

If Not r.ListObject Is Nothing Then
    IsDataFiltered = r.ListObject.AutoFilter.FilterMode
Else
    iRC = r.Rows.Count
    iFC = r.SpecialCells(xlCellTypeVisible).Count
    If iRC > iFC Then IsDataFiltered = True
End If

End Function