'Source: Chapter 26 - Excel VBA Notes for Professionals
'Purpose: Write error number, description and Erl to log file and return error text
'The sub below generates an error as an example - the line where the compiler attempts to divide 1 by 0
Public Sub ProcedureWithError()

Dim i As Integer, j As Integer
On Error GoTo LogErr
10 j = 1 / 0 ' raises an error

ExitHandler:
    Debug.Print "i=" & i
    Exit Sub
LogErr:
    'Change hard coded literals in description below as required
    MsgBox LogErrors("MyModule", "ProcedureWithError", Err), vbExclamation, "Error " & Err.Number
    Stop
    Resume Next

End Sub

'The below function logs the errors
Public Function LogErrors(ByVal sModule As String, ByVal sProc As String, Err As ErrObject) As String
    
Dim sLogFile As String: sLogFile = ThisWorkbook.Path & Application.PathSeparator & "LogErrors.txt"
Dim sLogTxt As String
Dim iFile As Long

' Create error text
sLogTxt = sModule & "|" & sProc & "|Erl " & Erl & "|Err " & Err.Number & "|" & Err.Description
On Error Resume Next
    iFile = FreeFile
    Open sLogFile For Append As iFile
    Print iFile, Format$(Now(), "yyyy-mm-dd hh:mm:ss "); sLogTxt

Print #iFile,
Close iFile

' Return error text
LogErrors = sLogTxt

End Function

'Additional Code to show log file
Sub ShowLogFile()

Dim sLogFile As String
sLogFile = ThisWorkbook.Path & Application.PathSeparator & "LogErrors.txt"

On Error GoTo LogErr
    Shell "notepad.exe " & sLogFile, vbNormalFocus

ExitHandler:
    On Error Resume Next
    Exit Sub

LogErr:
    MsgBox LogErrors("MyModule", "ShowLogFile", Err), vbExclamation, "Error No " & Err.Number
    Resume ExitHandler

End Sub