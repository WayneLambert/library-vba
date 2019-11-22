'Finds the nth word from a text string. The sPhrase and n must be passed from the value of an object
Function Find_nthWord(sPhrase As String, n As Integer) As String

Dim iCurrentPos As Long
Dim iCurrWordNo As Long

Find_nthWord = vbNullString
iCurrWordNo = 1

'Remove leading spaces
sPhrase = Trim$(sPhrase)

For iCurrentPos = 1 To Len(sPhrase)
    If (iCurrWordNo = n) Then
        Find_nthWord = Find_nthWord & Mid$(sPhrase, iCurrentPos, 1)
    End If

    If (Mid(sPhrase, iCurrentPos, 1) = " ") Then
        iCurrWordNo = iCurrWordNo + 1
    End If
Next iCurrentPos

'Remove the rightmost space
Find_nthWord = Trim$(Find_nthWord)

End Function