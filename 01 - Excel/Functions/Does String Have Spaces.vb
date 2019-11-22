'Tests to see whether the string has spaces
Function HasSpaces(ByVal s As String) As Boolean
    Const sSPACE_CHAR = " "
    If InStr(s, sSPACE_CHAR) > 0 Then HasSpaces = True Else HasSpaces = False
End Function