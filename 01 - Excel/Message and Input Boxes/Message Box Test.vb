Sub MsgBoxTest()

Dim QuestResp As VbMsgBoxResult

If Me.tbPassword1 = vbNullString Then
    QuestResp = MsgBox("Are you sure you would like to cut make these cuts without setting a password to access the workbooks?", _
    vbYesNoCancel + vbQuestion, "No Password Selected")
    Select Case QuestResp
        Case vbYes
            GoTo CutFiles
        Case vbNo
            Exit Sub
        Case vbCancel
            Exit Sub
    End Select
CutFiles:
End If

End Sub