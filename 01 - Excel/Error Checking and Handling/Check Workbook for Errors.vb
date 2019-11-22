'Checks for any errors of any type on all worksheets
'#N/A, #REF, #NAME, #VALUE, etc
Sub CheckForErrors()

Dim ws As Worksheet
Dim r As Range
Dim iErrors As Long

For Each ws In ThisWorkbook.Worksheets
    For Each r In ws.UsedRange
        If IsError(r.Value) Then iErrors = iErrors + 1
    Next r
Next ws

MsgBox iErrors & " errors found.", vbCritical, "Errors Found"

End Sub