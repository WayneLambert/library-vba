Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _ 
   (ByVal lpBuffer As String, nSize As Long) As Long 

Function ReturnComputerName() As String

Dim sBuffer As String * 255 
Dim iLen As Long 
Dim sCompName As String 

sCompName = vbNullString
iLen = GetComputerName(sBuffer, 255) 
iLen = InStr$(1, sBuffer, Chr(0)) 
If (iLen > 0) Then 
    sCompName = Left$(sBuffer, iLen - 1) 
Else 
    sCompName = sBuffer 
End If 

ReturnComputerName = UCase$(Trim$(sCompName))

End Function