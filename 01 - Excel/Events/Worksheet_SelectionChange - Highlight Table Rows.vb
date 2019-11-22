Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)

Application.ScreenUpdating = False

If Not Intersect(ActiveSheet.ListObjects("tblTRS").DataBodyRange, ActiveCell) Is Nothing Then
    Target.ListObject.DataBodyRange.Interior.ColorIndex = xlNone
    Intersect(ActiveCell.EntireRow, cnTRS.ListObjects("tblTRS").DataBodyRange).Interior.Color = rgb(255,255,153)
End If

Application.ScreenUpdating = True

End Sub