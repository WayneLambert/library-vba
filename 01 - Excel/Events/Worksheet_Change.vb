Private Sub Worksheet_Change(ByVal Target As Range)

'If user makes change to more than one cell, then exit procedure
If Target.Cells.Count > 1 Then Exit Sub

'If user makes an erroneous entry, then exit procedure
Is IsError(Target) Then Exit Sub

'If user deletes an entry, then exit procedure
If Target.Value = vbNullString Then Exit Sub

'Useful for just the one range intersection to test
If Not Intersect (Target,Me.Range("SelectedValue"))is Nothing Then
    'Insert code here - this code might include an If test to test whether the value in the range was changed to a certain value
End If

'A more condensed format for when there are multiple range intersections to test
Select Case True
	Case Not Intersect(Target,Me.Range("SelectedValue")) Is Nothing Then
    'Insert code here - this code might include an If test to test whether the value in the range was changed to a certain value
End Select

End Sub