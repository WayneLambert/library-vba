'Removes additional spaces between words
Function RemoveAdditionalSpaces(ByRef sStr As String) As String

Dim s As String

Do
    s = sStr
    sStr = Replace(sStr, Space(2), Space(1))
Loop Until s = sStr

RemoveAdditionalSpaces = sStr

End Function