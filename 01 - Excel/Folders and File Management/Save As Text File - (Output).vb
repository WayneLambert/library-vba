Sub SaveTextFile()
    Dim Filename As String, LineText As String
    Dim MyRange As Range, RowNo As Long, ColNo As Long
    Dim NoOfRows As Long, NoOfCols As Long
    
    Filename = ThisWorkbook.Path & "\textfile-" & Format(Now, "yyyy-mm-dd hhmmss") & ".txt"
    
    Open Filename For Output As #1
    
    Set MyRange = Range("Table2")
    NoOfRows = MyRange.Rows.Count
    NoOfCols = MyRange.Columns.Count - 1
    
    For RowNo = 2 To NoOfRows
        For ColNo = 1 To NoOfCols
            LineText = IIf(ColNo = 1, "", LineText & ",") & MyRange.Cells(RowNo, ColNo)
        Next ColNo
        Print #1, LineText
    Next RowNo
    
    Close #1
End Sub