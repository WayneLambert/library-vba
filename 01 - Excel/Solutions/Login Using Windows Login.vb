'To make this solution work...
'Add a reference to Microsoft Forms 2.0 Object Library
'Call the form = Login_Form
'Control names: username, password, domain, Login_Button, Cancel_login

'// In the form's module
Private Sub Cancel_Login_Click()
    Me.Hide
End Sub

'// In the form's module
Private Sub Login_Button_Click()
    Dim bLoginSuccess As Boolean
    
    bLoginSuccess = GetWindowsLogin(Me.UserName, Me.Password, Me.domain)
    MsgBox "Login " & bLoginSuccess
    Me.UserName = vbNullString
    Me.Password = vbNullString
    Me.Hide
End Sub

'// In the form's module
Private Sub UserForm_Activate()
    Me.UserName = Environ$("username")
    Me.domain = Environ$("userdomain")
End Sub

'// In a standard module
'Authenticates user and password entered with Active Directory
'Launch using VBA form - Login_Form (Username and Domain auto-populated)
Function GetWindowsLogin(ByVal sUserName As String, ByVal sPassword As String, ByVal sDomain As String) As Boolean

Dim oADsObject As Object, oADsNamespace As Object
Dim sADsPath As Object

On Error GoTo IncorrectPassword
    sADsPath = "WinNT://" & sDomain
    Set oADsObject = GetObject(sADsPath)
    Set oADsNamespace = GetObject("WinNT:")
    Set oADsObject = oADsNamespace.OpenDSObject(sADsPath, sDomain & "\" & sUserName, sPassword, 0)
On Error GoTo 0

GetWindowsLogin = True          'Access Granted

ExitHandler:
    Exit Function

IncorrectPassword:
    GetWindowsLogin = False     'Access Denied
    Resume ExitHandler

End Function