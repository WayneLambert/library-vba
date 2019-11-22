'Toggles the autofilter of the table object the opposite to what
'it currently is for each Excel table in Sheet1
Sub ToggleTableFilter()

For Each lo in Sheet1.ListObjects
    If Not lo.AutoFilter Is Nothing Then lo.Range.AutoFilter
Next lo

End Sub