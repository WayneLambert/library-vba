'// Call command to invoke No Of Positions checking sub //
'[Place inside code window for userform]

Private Sub tbNoOfPositions_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call mdlHiringManagerFormValidation.tbNoOfPositions_Exit(ByVal Cancel)
End Sub

**********************************************************************************

'[Insert code into a standard module]
Public Sub tbNoOfPositions_Exit(ByVal Cancel As MSForms.ReturnBoolean)

With frmR2Rs
    If Not IsNumeric(.tbNoOfPositions.Value) Then
        .tbNoOfPositions.BackColor = &HFF& 'Change the colour of the textbox to red if the field has not been populated
        MsgBox "You must enter a number between 1 and 100. You cannot leave the field blank.", vbCritical, "Value Required"
        Cancel = True 'Setting Cancel to True means the manager cannot leave this textbox until the value has been populated
    ElseIf .tbNoOfPositions.Value < 1 Then
        .tbNoOfPositions.BackColor = &HFF& 'Change the colour of the textbox to red if the number of vacancies is greater than 0
        MsgBox "You must enter a number greater than or equal to 1.", vbCritical, "Entry Not Permitted"
        Cancel = True  'Setting Cancel to True means the manager cannot leave this textbox until the value is greater than or equal to 1
    ElseIf .tbNoOfPositions.Value > 100 Then
        .tbNoOfPositions.BackColor = &HFF& 'Change the colour of the textbox to red if the number of vacancies is greater than 100
        MsgBox "This R2R template handles the creation of 100 positions.", vbCritical, "Entry Not Permitted"
        Cancel = True  'Setting Cancel to True means the manager cannot leave this textbox until the value is reduced to a number between 1-100
        .tbNoOfPositions.Value = vbNullString
    Else
        .tbNoOfPositions.BackColor = &H80000005 'Change colour of the textbox to show input is accepted
    End If
End With

End Sub