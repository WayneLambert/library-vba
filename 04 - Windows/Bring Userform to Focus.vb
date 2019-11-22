Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (
      ByVal lpClassName As String,
      ByVal lpWindowName As String
   ) As Long

Private Declare Function SetForegroundWindow Lib "user32" (
      ByVal hwnd As Long
   ) As Long

Public Sub ActivateUserForm()

    SetForegroundWindow DialogHWnd(UserForm1)

End Sub

Public Function DialogHWnd(
      ByRef WindowObject As Object
   ) As Long

    ' Return the hWnd value for the window.

    If TypeName(WindowObject) = "DialogSheet" Then
        Select Case (CDbl(Application.Version))
            Case 7 ' Excel 95
                DialogHWnd = GetWindowFromTitle(WindowObject.DialogFrame.Caption, "bosa_sdm_XL")
            Case 8 ' Excel 97
                DialogHWnd = GetWindowFromTitle(WindowObject.DialogFrame.Caption, "bosa_sdm_XL8")
            Case 9 ' Excel 2000
                DialogHWnd = GetWindowFromTitle(WindowObject.DialogFrame.Caption, "bosa_sdm_XL9")
            Case Else
                Exit Function
        End Select
    Else
        Select Case (CDbl(Application.Version))
            Case 8 ' Excel 97
                DialogHWnd = GetWindowFromTitle(WindowObject.Caption, "ThunderXFrame")
            Case Is >= 9 ' Excel 2000 or later
                DialogHWnd = GetWindowFromTitle(WindowObject.Caption, "ThunderDFrame")
            Case Else
                Exit Function
        End Select
    End If

End Function

Public Function GetWindowFromTitle(
      ByVal WindowTitle As String,
      Optional ByVal ClassName As String _
   ) As Long

    ' Find the window handle of the window with the class and name provided.

    Dim hwnd As Long

    If Len(ClassName) = 0 Then
        hwnd = FindWindow(vbNullString, WindowTitle)
    Else
        hwnd = FindWindow(ClassName, WindowTitle)
    End If

    GetWindowFromTitle = hwnd

End Function