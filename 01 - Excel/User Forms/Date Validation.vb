Public Sub tbStartDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim varStartDate As Date
Dim QuestionToMsgBox As String
Dim varYesNo As String

If IsDate(frmODFs.tbStartDate) Then
varStartDate = frmODFs.tbStartDate
Else: GoTo TestStartDateInput
End If

TestStartDateInput:
With frmODFs

    If Not IsDate(.tbStartDate) Then
        .tbStartDate.BackColor = &HFF&
            .tbStartDate.Value = vbNullString
            .tbStartDate.SelStart = 0
            .tbStartDate.SelLength = Len(tbToDate.Text) 'On error selects entire cells contents so it can be overtyped
            MsgBox "The value must be entered in a date format.", vbCritical, "Incorrect Date Format"
        Cancel = True
    ElseIf varStartDate < Date Then
        .tbStartDate.BackColor = &HFF&
        QuestionToMsgBox = "Are you sure you want to offer this role retrospectively?"
        varYesNo = MsgBox(QuestionToMsgBox, vbYesNo, "Confirmation Required")
            If varYesNo = vbYes Then
                GoTo StartDateAccept
            Else
                .tbStartDate.BackColor = &HFF&
                .tbStartDate = vbNullString
                Cancel = True
            End If
    ElseIf varStartDate > Date + 183 Then
        .tbStartDate.BackColor = &HFF&
        QuestionToMsgBox = "Are you sure you want to offer this role for a date greater than 6 months in advance?"
        varYesNo = MsgBox(QuestionToMsgBox, vbYesNo, "Confirmation Required")
            If varYesNo = vbYes Then
                GoTo StartDateAccept
                Else
                .tbStartDate.BackColor = &HFF&
                .tbStartDate = vbNullString
                Cancel = True
            End If
    Else
StartDateAccept:
        .tbStartDate.BackColor = &H80000005
        varStartDate = .tbStartDate.Value
        .tbStartDate = Format(varStartDate, "dd/mm/yy")
    End If
End With

End Sub