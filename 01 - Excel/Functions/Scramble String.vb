's - The number or text you want to scramble
Function ScrambleString(s As String) As String

Dim iLen As Integer, i As Integer, iRandPos As Integer
Dim sChar As String * 1

iLen = Len(s)
For i = 1 To iLen
    sChar = Mid$(s, i, 1)
    iRandPos = Int((iLen - 1 + 1) * Rnd + 1)
    Mid$(s, i, 1) = Mid$(s, iRandPos, 1)
    Mid$(s, iRandPos, 1) = sChar
Next i
    
ScrambleString = s

End Function