Private Sub btnOK_Click()

    Dim Ctrl As Control
    Dim CtrlType1 As String
    Dim SelectedOption As String

    CtrlType1 = "OptionButton"

    For Each Ctrl In UserForm1.Controls
        If TypeName(Ctrl) = CtrlType1 And Ctrl.Value = True Then
            SelectedOption = Ctrl.Caption
        End If
    Next Ctrl

    Sheet1.Cells(1, 1) = SelectedOption

End Sub