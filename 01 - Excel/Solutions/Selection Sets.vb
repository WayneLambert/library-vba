'I may not need to select the range. If the row or column is required, then use .row or .column

'Select Down (As in Ctrl + Shift + Down)
Sub SelectDown()
    Range(ActiveCell, ActiveCell.End(xlDown)).Select
End Sub

'Select Up (As in Ctrl + Shift + Up)
Sub SelectUp()
    Range(ActiveCell, ActiveCell.End(xlUp)).Select
End Sub

'Select Right (As in Ctrl + Shift + Right)
Sub SelectToRight()
    Range(ActiveCell, ActiveCell.End(xlToRight)).Select
End Sub

'Select Left (As in Ctrl + Shift + Left)
Sub SelectToLeft()
    Range(ActiveCell, ActiveCell.End(xlToLeft)).Select
End Sub

'Select Current Region (As in Ctrl + Shift + *)
Sub SelectCurrentRegion()
    ActiveCell.CurrentRegion.Select
End Sub

'Select Active Area (As in Ctrl + Shift + Home)
Sub SelectActiveArea()
    Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).Select
End Sub

'Select Contiguous Cells in ActiveCell's Column
Sub SelectActiveColumn()
    Dim TopCell As Range
    Dim BottomCell As Range

    If IsEmpty(ActiveCell) Then Exit Sub
'   ignore error if activecell is in Row 1
    On Error Resume Next
    If IsEmpty(ActiveCell.Offset(-1, 0)) Then Set TopCell = ActiveCell Else Set TopCell = ActiveCell.End(xlUp)
    If IsEmpty(ActiveCell.Offset(1, 0)) Then Set BottomCell = ActiveCell Else Set BottomCell = ActiveCell.End(xlDown)
    Range(TopCell, BottomCell).Select
End Sub

'Select Contiguous Cells in ActiveCell's Row
Sub SelectActiveRow()
    Dim LeftCell As Range
    Dim RightCell As Range
    
    If IsEmpty(ActiveCell) Then Exit Sub
'   ignore error if activecell is in Column A
    On Error Resume Next
    If IsEmpty(ActiveCell.Offset(0, -1)) Then Set LeftCell = ActiveCell Else Set LeftCell = ActiveCell.End(xlToLeft)
    If IsEmpty(ActiveCell.Offset(0, 1)) Then Set RightCell = ActiveCell Else Set RightCell = ActiveCell.End(xlToRight)
    Range(LeftCell, RightCell).Select
End Sub

'Select an Entire Column (As In Ctrl+Spacebar)
Sub SelectEntireColumn()
    ActiveCell.EntireColumn.Select
End Sub

'Select an Entire Row  (As In Shift+Spacebar)
Sub SelectEntireRow()
    ActiveCell.EntireRow.Select
End Sub

'Select the Entire Worksheet (As In Ctrl+A)
Sub SelectEntireSheet()
    Cells.Select
End Sub

'Activate the Next Blank Cell Below
Sub ActivateNextBlankDown()
    ActiveCell.Offset(1, 0).Select
    Do While Not IsEmpty(ActiveCell)
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub

'Activate the Next Blank Cell To the Right
Sub ActivateNextBlankToRight()
    ActiveCell.Offset(0, 1).Select
    Do While Not IsEmpty(ActiveCell)
        ActiveCell.Offset(0, 1).Select
    Loop
End Sub

'Select From the First NonBlank to the Last Nonblank in the Row
Sub SelectFirstToLastInRow()
    Dim LeftCell As Range
    Dim RightCell As Range
    
    Set LeftCell = Cells(ActiveCell.Row, 1)
    Set RightCell = Cells(ActiveCell.Row, 16384)

    If IsEmpty(LeftCell) Then Set LeftCell = LeftCell.End(xlToRight)
    If IsEmpty(RightCell) Then Set RightCell = RightCell.End(xlToLeft)
    If LeftCell.Column = 16384 And RightCell.Column = 1 Then ActiveCell.Select Else Range(LeftCell, RightCell).Select
End Sub

'Select From the First NonBlank to the Last Nonblank in the Column
Sub SelectFirstToLastInColumn()
    Dim TopCell As Range
    Dim BottomCell As Range
    
    Set TopCell = Cells(1, ActiveCell.Column)
    Set BottomCell = Cells(1048576, ActiveCell.Column)

    If IsEmpty(TopCell) Then Set TopCell = TopCell.End(xlDown)
    If IsEmpty(BottomCell) Then Set BottomCell = BottomCell.End(xlUp)
    If TopCell.Row = 1048576 And BottomCell.Row = 1 Then ActiveCell.Select Else Range(TopCell, BottomCell).Select
End Sub