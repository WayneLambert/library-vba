'sText1 - The text string containing your substring.
'sText2 - The text string to look for.
'bIgnoreCase - (Optional) Whether you want to ignore the case of the text.

Public Function CONTAINS(ByVal sText1 As String, _ 
                         ByVal sText2 As String, _ 
                Optional ByVal bIgnoreCase As Boolean = False) As Boolean 
                
Call Application.Volatile(True)
            
CONTAINS = False
If Len(sText1) = 0 Or Len(sText2) = 0 Then Exit Function

If bIgnoreCase = False Then
    If UCase$(sText2) Like "*" & UCase$(sText1) & "*" Then CONTAINS = True
Else 
    If sText2 Like "*" & sText1 & "*" Then CONTAINS = True
End If

End Function