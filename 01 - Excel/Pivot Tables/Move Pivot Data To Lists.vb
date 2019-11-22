Sub MovePivotDataToLists()

Dim PivotFilter As String
Dim x As Integer

For x = 1 To 56
    PivotFilter = cnTables.Range("SelectedPasteLocation").Offset(0, x - 1).Value
    cnProfAndTitles.Activate
    With ActiveSheet.PivotTables("PivotTable")
        .PivotFields("PRF Profession").ClearAllFilters
        .PivotFields("PRF Profession").CurrentPage = PivotFilter
        .PivotSelect "'PRF Role Title'[All]", xlLabelOnly, True
    End With
    Selection.Copy
    cnTables.Select
    Range("SelectedPasteLocation").Offset(1, x - 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Next x

End Sub