'Returns TRUE if the workbook is open
Private Function WorkbookIsOpen(wbname) As Boolean

Dim x As Workbook
On Error Resume Next
Set x = Workbooks(wbname)

If Err = 0 Then WorkbookIsOpen = True Else WorkbookIsOpen = False

End Function