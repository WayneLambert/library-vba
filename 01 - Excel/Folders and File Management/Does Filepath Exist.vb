'Returns TRUE if the path exists
Function PathExists(sPath) as Boolean
    On Error Resume Next
    PathExists = (GetAttr(sPath) and vbDirectory)= vbDirectory
End Function

'*****************************************************************************

'A slightly more explicit variation on the above...
'Again, returns TRUE if the path exists
Private Function PathExists(sPath) As Boolean
	If Dir(sPath, vbDirectory) = vbNullString Then
		PathExists = False
	Else
		PathExists = (GetAttr(sPath) And vbDirectory) = vbDirectory
	End If
End Function

'*****************************************************************************

'And a third variation using a reference to Microsoft Scripting Runtime library

Function PathExists(ByVal sPath As String) As Boolean

Dim fso As Scripting.FileSystemObject
Set fso = New Scripting.FileSystemObject
PathExists = fso.FolderExists(sPath)

End Function