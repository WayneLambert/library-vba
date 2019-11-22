'Check to see if FolderExists
Function FolderExists(sPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(sPath) And vbDirectory) = vbDirectory)
    On Error Goto 0
End Function

'Use the TrailingSlash() function to add a slash to the end of a path unless it is already there.
Function TrailingSlash(v As Variant) As String
    If Len(v) > 0 Then
        If Right(v, 1) = "\" Then TrailingSlash = v Else TrailingSlash = v & "\"
    End If
End Function