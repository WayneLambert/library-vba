'Deletes all data rows from an Excel table (except the first row)
Sub ClearDownTable()

Dim tbl As ListObject
Set tbl = ActiveSheet.ListObjects("Table1")

'Delete all table rows except first row
With tbl.DataBodyRange
    If .Rows.Count > 1 Then
        .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).Rows.Delete
    End If
End With

'Clear out data from first table row
tbl.DataBodyRange.Rows(1).ClearContents

'If you would like to retain the formulas in the table, use this row instead of the one above
tbl.DataBodyRange.Rows(1).SpecialCells(xlCellTypeConstants).ClearContents

End Sub