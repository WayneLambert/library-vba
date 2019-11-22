'Use the procedure below if you would like to register the function
Sub RegisterUDF()

Dim sFunc As String     'name of the function you want to register
Dim sDesc As String     'description of the function itself
Dim sArgs() As String   'description of function arguments

'Register SpellNumber function
ReDim sArgs(1 To 2)     'Upper bound of array is the number of arguments you would like to register
sFunc = "SpellNumber"
sDesc = "Converts a numeric field entry into English language words"
sArgs(1) = "The range field that contains a numeric entry"
sArgs(2) = "Enter TRUE if you would like UPPERCASE, else FALSE or leave blank for lowercase"
Application.MacroOptions Macro:=sFunc, Description:=sDesc, ArgumentDescriptions:=sArgs, Category:="My Functions"

End Sub

'Converts a numeric field entry into English language words
Public Function SpellNumber(ByVal n As Variant, Optional ByVal UPPERCASE As Boolean) As String

Dim Integers As Variant
Dim sFirstDP As String, sSecondDP As String, sTmp As String, sTidyUp As String
Dim bNegativeNum As Boolean
Dim iDecPlace As Long, iCount As Long

ReDim Place(9) As String
Place(2) = " Thousand, "
Place(3) = " Million, "
Place(4) = " Billion, "
Place(5) = " Trillion, "

'Test if negative
bNegativeNum = IIf(n < 0, True, False)

'String representation of amount.
n = Trim$(Str(Abs(Round(n, 2))))

'Position of decimal place 0 if none
iDecPlace = InStr(n, ".")

'Convert sFirstDP and set n to amount
If iDecPlace > 0 Then
    sFirstDP = GetDigit(Mid$(n, iDecPlace + 1, 1))
    sSecondDP = GetDigit(Mid$(n, iDecPlace + 2, 1))
    n = Trim$(Left$(n, iDecPlace - 1))
End If

iCount = 1
Do While n <> vbNullString
    sTmp = GetHundreds(Right$(n, 3))
    If sTmp <> vbNullString Then Integers = sTmp & Place(iCount) & Integers
    If Len(n) > 3 Then n = Left$(n, Len(n) - 3) Else n = vbNullString
    iCount = iCount + 1
Loop

Select Case Integers
    Case vbNullString
        Integers = "Zero Point"
    Case "One"
        Integers = "One"
     Case Else
        Integers = Integers
End Select

Select Case sFirstDP
    Case vbNullString
        sFirstDP = vbNullString
    Case "One"
        sFirstDP = " zero one"
    Case Else
        sFirstDP = " point " & sFirstDP
End Select

'Tidy up any additional characters
sTidyUp = Integers & sFirstDP & " " & sSecondDP
sTidyUp = IIf(Right$(sTidyUp, 3) = ",  ", Replace(sTidyUp, ",  ", ""), sTidyUp)
sTidyUp = IIf(Right$(sTidyUp, 6) = " and  ", Replace(sTidyUp, " and  ", ""), sTidyUp)

If bNegativeNum = True Then sTidyUp = "Minus " & Trim$(sTidyUp)
If UPPERCASE = True Then SpellNumber = UCase$(sTidyUp) Else SpellNumber = sTidyUp

End Function

'*********************************************************************************************************************

'Converts a number from 100-999 into text
Function GetHundreds(ByVal n As String)

Dim sResult As String

If Val(n) = 0 Then Exit Function

n = Right("000" & n, 3)

'Convert the hundreds place
If Mid$(n, 1, 1) <> "0" Then sResult = GetDigit(Mid$(n, 1, 1)) & " Hundred and "

'Convert the tens and ones place
If Mid$(n, 2, 1) <> "0" Then sResult = sResult & GetTens(Mid$(n, 2)) Else sResult = sResult & GetDigit(Mid$(n, 3))

GetHundreds = sResult

End Function

'*********************************************************************************************************************

'Converts a number from 10 to 99 into text.
Function GetTens(ByVal sTensText As String)

Dim sResult As String
sResult = vbNullString                                          'Null out the temporary function value

If Val(Left$(sTensText, 1)) = 1 Then                            'If value between 10-19...
    Select Case Val(sTensText)
        Case 10: sResult = "Ten"
        Case 11: sResult = "Eleven"
        Case 12: sResult = "Twelve"
        Case 13: sResult = "Thirteen"
        Case 14: sResult = "Fourteen"
        Case 15: sResult = "Fifteen"
        Case 16: sResult = "Sixteen"
        Case 17: sResult = "Seventeen"
        Case 18: sResult = "Eighteen"
        Case 19: sResult = "Nineteen"
        Case Else
    End Select
Else                                                            'If value between 20-99...
    Select Case Val(Left(sTensText, 1))
        Case 2: sResult = "Twenty "
        Case 3: sResult = "Thirty "
        Case 4: sResult = "Forty "
        Case 5: sResult = "Fifty "
        Case 6: sResult = "Sixty "
        Case 7: sResult = "Seventy "
        Case 8: sResult = "Eighty "
        Case 9: sResult = "Ninety "
        Case Else
    End Select
    sResult = sResult & GetDigit(Right$(sTensText, 1))           'Retrieve ones place
End If

GetTens = sResult

End Function

'*********************************************************************************************************************

'Converts a number from 1 to 9 into text.
Function GetDigit(ByVal Digit As String)

Select Case Val(Digit)
    Case 1: GetDigit = "One"
    Case 2: GetDigit = "Two"
    Case 3: GetDigit = "Three"
    Case 4: GetDigit = "Four"
    Case 5: GetDigit = "Five"
    Case 6: GetDigit = "Six"
    Case 7: GetDigit = "Seven"
    Case 8: GetDigit = "Eight"
    Case 9: GetDigit = "Nine"
    Case Else: GetDigit = vbNullString
End Select

End Function