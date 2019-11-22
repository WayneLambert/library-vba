'Launches and terminates the calculator as necessary
Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, _
    lpExitCode As Long) As Long

'This is the most basic implementation. Probably no need for the other procedures
Sub StartCalc()

Dim Program As String
Dim TaskID As Double

On Error Resume Next
Program = "calc.exe"
TaskID = Shell(Program, 1)
If Err <> 0 Then MsgBox "Cannot start " & Program, vbCritical, "Error"

End Sub

Sub StartCalc2()

Const ACCESS_TYPE As Integer = &H400
Const STILL_ACTIVE As Integer = &H103
Dim TaskID As Long
Dim hProc As Long
Dim iExitCode As Long
Dim Program As String

Program = "calc.exe"
On Error Resume Next

'Shell the task
TaskID = Shell(Program, 1)

'Get the process handle
hProc = OpenProcess(ACCESS_TYPE, False, TaskID)

If Err <> 0 Then
    MsgBox "Cannot start " & Program, vbCritical, "Error"
    Exit Sub
End If

Do
    GetExitCodeProcess hProc, iExitCode
    DoEvents
Loop While iExitCode = STILL_ACTIVE

'Task is finsihed, so show message
MsgBox Program & " was closed."

End Sub

Sub ActivateCalc()

Dim sAppFile As String
Dim dCalcTaskID As Double

sAppFile = "Calc.exe"
On Error Resume Next
AppActivate "Calculator"

If Err <> 0 Then
    Err = 0
    dCalcTaskID = Shell(sAppFile, 1)
    If Err <> 0 Then MsgBox "Can't start Calculator"
End If

End Sub