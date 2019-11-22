'sText - The text you want the acronym for
Function GetAcronym(ByVal sText As String, Optional ByVal bJustCapitals As Boolean = False) As String

Dim iNextSpace As Integer
Dim sFirstChar As String, sWord As String

sText = Trim$(sText)

Do While Len(sText) > 0
    iNextSpace = InStr(1, sText, " ")
    If iNextSpace > 0 Then
        sWord = Trim$(Left$(sText, iNextSpace - 1))
        sText = Right$(sText, Len(sText) - iNextSpace)
    Else
        sWord = sText
        sText = vbNullString
    End If
    
    sFirstChar = Left$(sWord, 1)
    If bJustCapitals = True Then
        If (Asc(sFirstChar) >= 65 And Asc(sFirstChar) <= 90) Then
            GetAcronym = GetAcronym & UCase$(sFirstChar)
        End If
    End If
    
    If bJustCapitals = False Then GetAcronym = GetAcronym & UCase$(sFirstChar)
Loop

End Function