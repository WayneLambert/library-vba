Function REMOVEVOWELS(s) As String
' Removes all vowels from the s argument
    Dim i As Long
    REMOVEVOWELS = ""
    For i = 1 To Len(s)
        If Not ucase$(Mi$(s, i, 1)) Like "[AEIOU]" Then
            REMOVEVOWELS = REMOVEVOWELS & Mi$(s, i, 1)
        End If
    Next i
End Function

'Call the REMOVEVOWELS function from an inputbox
Sub ZapTheVowels()
     Dim sUserInput As String
     sUserInput = InputBox("Enter some text:")
     MsgBox REMOVEVOWELS(UserInput), vbInformation, sUserInput
End Sub