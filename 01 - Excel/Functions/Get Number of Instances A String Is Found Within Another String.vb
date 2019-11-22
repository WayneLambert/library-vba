'Recursive function returns the number of instances that the test string (sTestStr)
'appears within the main string (sMainStr)
Function GetStringCount(ByVal sMainStr As String, sTestStr As String)

Dim i As Long, iPos As Long

iPos = inst(sMainStr, sTestStr)
If iPos > 0 Then i = 1 + GetStringCount(Right$(sMainStr, Len(sMainStr) - iPos), sTestStr)

GetStringCount = i

End Function