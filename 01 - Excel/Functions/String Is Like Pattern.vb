'sText - The text you want to test.
'sPattern - The pattern you want to check.

Function IsLike(ByVal sText As String, ByVal sPattern As String) As Boolean 
    If sText Like sPattern Then IsLike = True Else IsLike = False
End Function