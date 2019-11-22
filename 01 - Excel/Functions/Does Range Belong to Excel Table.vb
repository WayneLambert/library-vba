'Determines whether the range passed through belongs to an Excel Table (ListObject)
Function IsExcelTable(ByRef r As Range) As Boolean
    If Not r.ListObject Is Nothing Then IsExcelTable = True Else IsExcelTable = False
End Function