
#If VBA7 And Win64 Then
  Declare PtrSafe Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#Else
  Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#End If

'Return the Windows directory to a message box
Sub ShowWindowsDir()
    Dim WinPath As String * 255
    Dim WinDir As String
    WinPath = Space(255)
    WinDir = Left(WinPath, GetWindowsDirectoryA(WinPath, Len(WinPath)))
    MsgBox WinDir, vbInformation, "Windows Directory"
End Sub

'Return the windows directory as a value; accessible to the worksheet
Function WINDOWSDIR() As String
'   Returns the Windows directory
    Dim WinPath As String * 255
    WinPath = Space(255)
    WINDOWSDIR = Left(WinPath, GetWindowsDirectoryA(WinPath, Len(WinPath)))
End Function