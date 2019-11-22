
Private Sub Worksheet_Change(ByVal Target As Range)

Dim r As Range

Application.EnableEvents = False

For Each r In Target
    If Not Application.Intersect(r, Range("hcNamedRange")) Is Nothing Then
        If Not IsNumeric(r.Value) Then
            r.Value = vbNullString
        End If
    End If
Next r

Application.EnableEvents = True

End Sub