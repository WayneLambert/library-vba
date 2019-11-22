'Builds out a message box report in memory and returns to the user
Sub MsgBoxReport()

Dim msg as string
Dim Variable1 as string, Variable2 as string, Variable3 as string, Variable4 as string

    Msg = "Variable 1 Ans:" & vbTab & Variable1 & vbCrLf
    Msg = Msg & "Variable 2 Ans:" & vbTab & Variable1 & vbCrLf
    Msg = Msg & "Variable 3 Ans: " & vbTab & Variable1 & vbCrLf
    Msg = Msg & "Variable 4 Ans: " & vbTab & Variable1 & vbCrLf
    MsgBox Msg, vbInformation, "Report"
End Sub