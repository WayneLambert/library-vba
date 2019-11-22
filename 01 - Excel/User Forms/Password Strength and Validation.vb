'Insert within the userform
Private Sub tbPassword1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

Dim Pwd As String

Pwd = Me.tbPassword1
    Select Case HasSpaces(Pwd)
        Case Is = True
            MsgBox "The password contains spaces. Please retry using a password without spaces.", vbCritical, "Password Contains Spaces..."
            frmCutting.tbPassword1 = vbNullString
            Cancel = True
            Exit Sub
    End Select
    
    Select Case PasswordStrengthCheck(Pwd)
        Case Is = "Weak"
            MsgBox "The password is too weak. Please ensure that it is 8 characters long and contains UPPERCASE, lowercase, numbers and special characters.", vbCritical, "Password Too Weak..."
            frmCutting.tbPassword1 = vbNullString
            Cancel = True
            Exit Sub
        Case Is = "Medium"
            MsgBox "The password is too weak. Please ensure that it is 8 characters long and contains UPPERCASE, lowercase, numbers and special characters.", vbCritical, "Password Too Weak..."
            frmCutting.tbPassword1 = vbNullString
            Cancel = True
            Exit Sub
        Case Is = "Strong"
            Exit Sub
    End Select

End Sub

'Insert within a standard code module
Function HasSpaces(ByVal Pwd As String) As Boolean

Const SPACE_CHAR As String = " "

'Check for spaces
If InStr(Pwd, SPACE_CHAR) > 0 Then
    HasSpaces = True
    Exit Function
Else: HasSpaces = False
End If

End Function

'Insert within a standard code module
Function PasswordStrengthCheck(ByVal Pwd As String) As String

Dim Strength As Long, PasswordLength As Long
Dim i As Long
Dim HasLcase As Boolean, HasUcase As Boolean, HasNo As Boolean
Dim HasSpecChar1 As Boolean, HasSpecChar2 As Boolean, HasSpecChar3 As Boolean, HasSpecChar4 As Boolean

PasswordLength = Len(Pwd)

'Check length
Select Case PasswordLength
    Case Is >= 8
        Strength = Strength + 1
    Case Is >= 4
        Strength = 1
    Case Else
        GoTo CalculateStrength
End Select

'Loop through each character in the password submitted
For i = 1 To PasswordLength
    Select Case Asc(Mid$(Pwd, i, 1))
        Case 33 To 47       'special characters
            If Not HasSpecChar1 Then
                Strength = Strength + 1
                HasSpecChar1 = True
            End If
        Case 48 To 57       'numbers
            If Not HasNo Then
                Strength = Strength + 1
                HasNo = True
            End If
        Case 58 To 64       'special characters
            If Not HasSpecChar2 Then
                Strength = Strength + 1
                HasSpecChar2 = True
            End If
        Case 65 To 90       'UPPERCASE
            If Not HasUcase Then
                Strength = Strength + 1
                HasUcase = True
            End If
        Case 91 To 96       'special characters
            If Not HasSpecChar3 Then
                Strength = Strength + 1
                HasSpecChar3 = True
            End If
        Case 97 To 122      'lowercase
            If Not HasLcase Then
                Strength = Strength + 1
                HasLcase = True
            End If
        Case 123 To 255     'special characters
            If Not HasSpecChar4 Then
                Strength = Strength + 1
                HasSpecChar4 = True
            End If
    End Select
Next i

CalculateStrength:
    Select Case Strength
        Case 0 To 2
            PasswordStrengthCheck = "Weak"
        Case 3
            PasswordStrengthCheck = "Medium"
        Case Is >= 4
            PasswordStrengthCheck = "Strong"
    End Select

If PasswordLength < 8 And PasswordStrengthCheck = "Strong" _
Then PasswordStrengthCheck = "Medium"

End Function