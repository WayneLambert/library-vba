'Returns the filename from a path/filename string
Private Function FileNameOnly(sPathName) As String

Dim vTmp As Variant

length = Len(sPathName)
vTmp = Split(sPathName, Application.PathSeparator)
FileNameOnly = vTmp(UBound(vTmp))

End Function