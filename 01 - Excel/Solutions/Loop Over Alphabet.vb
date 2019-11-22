Sub LoopOverAlphabet()

Const iNO_OF_LETTERS as Integer = 26     'number of letters you want to loop through - declared as constant as the English alphabet always has 26 letters
Dim sFirstLetter As String * 1, s As String * 1
Dim bCapitalLetters As Boolean
Dim iChr As Long, i As Long

sFirstLetter = "A"        'the letter you want to start with
bCapitalLetters = True    'set to True if you want capital letters (A B C). False if you want lowercase (a b c)

If bCapitalLetters = True Then sFirstLetter = UCase$(sFirstLetter)
If bCapitalLetters = False Then sFirstLetter = LCase$(sFirstLetter)
iChr = Asc(sFirstLetter)

For i = 1 To iNO_OF_LETTERS
    If bCapitalLetters = True Then
        If iChr > 90 Then iChr = 64 + iChr - 90
        s = Chr(iChr)
    Else
        If iChr > 122 Then iChr = 96 + iChr - 122
        s = Chr(iChr)
    End If
    iChr = iChr + 1
    Debug.Print s
Next i

End Sub