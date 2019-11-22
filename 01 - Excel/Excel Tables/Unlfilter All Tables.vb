'Unfilters all of the Excel tables within each worksheet of the workbook
Sub UnfilterAllTables()

Dim ws as worksheet
Dim lo as ListObject, los as ListObjects

For Each ws in ThisWorkbook.Worksheets
    For Each lo in ws.ListObjects
        lo.AutoFilter.ShowAllData
    Next lo  
Next ws

End Sub