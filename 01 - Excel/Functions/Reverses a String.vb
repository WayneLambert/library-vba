'Returns a string passed in in reverse order. e.g. "on" becomes "no"
Function ReverseString(ByVal sOrig As String) As String
 
Dim sReversed As String, s As String
Dim i As Long
 
For i = Len(sOrig) To 1 Step -1
    s = Mid$(sOrig, i, 1)
    sReversed = sReversed & s
Next i

ReverseString = sReversed
 
End Function

'Method 2
's - The number or text you want to reverse
Function ReverseString(ByVal s As String) As String 

   If Application.WorksheetFunction.IsNonText(s) = True Then 
      ReverseString = CVErr(xlCVError.xlErrNA) 
   Else 
      ReverseString = StrReverse(s) 
   End If 

End Function