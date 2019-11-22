Function GetCharacterType(ByVal sCharacter As String) As CharType

Select Case AscW(sCharacter)
    Case Is < 0
        sCharacterType = Unicode
    Case Is > 255
        sCharacterType = Unicode
    Case 48 To 57
        sCharacterType = Number
    Case 65 To 90
        sCharacterType = UpperCase
    Case 97 To 122
        sCharacterType = LowerCase
    Case Else
        sCharacterType = Other
End Select

End Function