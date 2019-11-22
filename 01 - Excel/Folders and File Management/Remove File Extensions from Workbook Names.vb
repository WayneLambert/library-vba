'Uses an MS Access listbox to populate a list of workbooks which include their file extensions
'Custom iUbound variable is used to set the fixed array to the same size as the number of items in the list (The -1 is because the array is at base 0)
Sub RemoveFileExtensionsFromWorkbookNames()

Dim Exts() As Variant
Dim Workbooks() As String
Dim c As Variant
Dim iUbound As Long, i As Long

iUbound = Form_frmCuttingOptions.lbWorkbooks.ListCount - 1
ReDim Workbooks(0 To iUbound) As String

For i = 0 To iUbound
    Workbooks(i) = Form_frmCuttingOptions.lbWorkbooks.ItemData(i)
    Exts = Array(".xlsx", ".xlsm", ".xlsb", ".xls")
    For Each c In Exts
        Workbooks(i) = Replace(Workbooks(i), c, "")
    Next c
Next i

Erase Exts

End Sub
