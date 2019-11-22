'If your project works with files other than Excel files, you should test whether a file is already open by another process before you attempt to read it or write to it. This page describes a function named IsFileOpen that returns True if the specified file is open or returns False if the specified file is not open. The code works by simply attempting to open the file for exclusive access. If the file is open by another process, the attempt to open it will fail. If the file is not in use, the attempt to open it will succeed. Once opened, the file is immediately closed without saving.

Public Function IsFileOpen(FileName As String, Optional ResultOnBadFile As Variant) As Variant
 
Dim FileNum As Integer, ErrNum As Integer
Dim V As Variant

On Error Resume Next
 
If Trim(FileName) = vbNullString Then
    If IsMissing(ResultOnBadFile) = True Then
        IsFileOpen = False
    Else
        IsFileOpen = ResultOnBadFile
    End If
    Exit Function
End If
 
V = Dir(FileName, vbNormal)
If IsError(V) = True Then
    ' syntactically bad file name
    If IsMissing(ResultOnBadFile) = True Then
        IsFileOpen = False
    Else
        IsFileOpen = ResultOnBadFile
    End If
    Exit Function
ElseIf V = vbNullString Then
    ' file doesn't exist.
    If IsMissing(ResultOnBadFile) = True Then
        IsFileOpen = False
    Else
        IsFileOpen = ResultOnBadFile
    End If
    Exit Function
End If

FileNum = FreeFile()
 
Err.Clear
Open FileName For Input Lock Read As #FileNum
ErrNum = Err.Number
 
Close FileNum
On Error GoTo 0
 
Select Case ErrNum
    Case 0
         
        IsFileOpen = False
    Case 70
         
        IsFileOpen = True
    Case Else
         
        IsFileOpen = True
End Select

End Function