Private Sub Worksheet_Calculate()

Static OldValue

If Range("SelectedValue").Value = vbNullString Then Exit Sub

If Range("SelectedValue").Value <> OldUBR Then
    OldUBR = Range ("SelectedValue").Value
    Call SetSelectedValue
End If

End Sub