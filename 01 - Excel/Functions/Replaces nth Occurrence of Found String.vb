'Example of how to call the Replace nth Occurrence function
'Replaces the 3rd instance of "plane" with "boat"
Sub ReplaceString()

sToReplace = "One plane, two plane, red plane, blue plane"
sToReplace = ReplaceNthOccurrence(sToReplace, "plane", "boat", 3)

End Sub

'Replaces the nth occurrence of a substring in a string
'Parameters are s1:=The string to do the replacements on, sFind:=What do you want to replace?
'sReplace:=What would you like to replace it with?, iCount:=Which occurrence of it being found would you like to replace
Function ReplaceNthOccurrence(ByVal s1 As Variant, sFind As String, _
    sReplace As String, n As Long, Optional iCount As Long) As String

Dim sM As String: sM = s1
Dim i As Long, j As Long

If iCount <= 0 Then iCount = 1

For i = 1 To n - 1
    j = InStr(1, sM, sFind)
    sM = Mid$(sM, j + Len(sFind), Len(sM))
Next i

If n <= 0 Then
    ReplaceNthOccurrence = s1
Else
    ReplaceNthOccurrence = Mid$(s1, 1, Len(s1) - Len(sM)) & Replace$(sM, sFind, sReplace, Start:=1, Count:=iCount)
End If

End Function