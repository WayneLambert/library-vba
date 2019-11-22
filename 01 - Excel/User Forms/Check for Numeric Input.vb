Public Sub tbNoOfPositions_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

With frmR2Rs
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0 ' this prevents non-numeric data from showing up in the TextBox form field
        MsgBox "Please enter a number. Text entries are not permitted.", vbCritical , "Entry Not Permitted."
    Else
        .tbNoOfPositions.BackColor = &H80000005 'Change colour of the textbox to show input is accepted
    End If
End With

End Sub