Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)

Cells.Interior.ColorIndex = xlNone
With ActiveCell
    .EntireRow.Interior.Color = RGB(219, 229, 241)
    .EntireColumn.Interior.Color = RGB(219, 229, 241)
End With

End Sub

Private Sub Worksheet_BeforeRightClick _
  (ByVal Target As Excel.Range, Cancel As Boolean)
    Cancel = True
    MsgBox "The shortcut menu is not available."
End Sub