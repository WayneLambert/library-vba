Dim NextTick As Date

Sub UpdateClock()

'Updates cell A1 with the current time
ThisWorkbook.Sheets(1).Range("A1") = Time
'Set up the next event five seconds from now
NextTick = Now + TimeValue("00:00:01")
Application.OnTime NextTick, "UpdateClock"

End Sub

'Enables a clock to be displayed within a label on a userform
Private Sub UserForm_Activate()

Dim CM As Boolean

Do
    If CM = True Then Exit Sub
    lblClock.Caption = Format(Now, "hh:mm:ss")
    DoEvents
Loop

End Sub

Sub StopClock()

'Cancels the OnTime event (stops the clock)
On Error Resume Next
Application.OnTime NextTick, "UpdateClock", , False

End Sub