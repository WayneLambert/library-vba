Public Sub Example1()

Dim Ctrl As Control
Dim CtrlType1 As String
Dim CtrlType2 As String

'Define types of controls on page of R2R form
CtrlType1 = "TextBox"
CtrlType2 = "ComboxBox"
CtrlType3 = "CheckBox"

'Loop Through each control on the Hiring Manager page of the R2R form
For Each Ctrl In Userform.Controls
    'Narrow down to specific type
    If TypeName(Ctrl) = CtrlType1 Or TypeName(Ctrl) = CtrlType2 Or TypeName(Ctrl) = CtrlType3 Then
        'Enable the form controls
        Ctrl.Enabled = True
    End If
Next Ctrl

End Sub

Public Sub Example2()

Dim r As range
Dim Ctrl As MSFrms.Control

Set r = ws.columns("A").row("6")

If r Is Nothing Then
    For Each Ctrl In MSForms.Textbox Then
        If TypeOf Ctrl Is msforms.textbox Then
            Ctrl.value = vbNullString
        End If
        Ctrl.value = vbNullString
    Next Ctrl
End If

End Sub