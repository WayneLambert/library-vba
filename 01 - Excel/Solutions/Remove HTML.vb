Sub HTML_Removal()

Dim Cell As Range

With CreateObject("vbscript.regexp")
    .Pattern = "\<.*?\>"
    .Global = True
    For Each Cell In Selection
        Cell.Value = .Replace(Cell.Value, "")
    Next
End With

End Sub