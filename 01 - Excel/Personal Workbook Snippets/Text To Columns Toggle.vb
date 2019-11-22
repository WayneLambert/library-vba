Option Explicit

'Sub toggles between General and Text datatypes
'Insert code into the Personal Macro Workbook. Add button to custom toolbar for quick access
Public Sub TextToColumnsToggle()

Dim ws As Worksheet
Dim ParseDataRange As Range, aCell As Range
Dim ActiveTable As ListObject
Dim StartingPosition As String, TableName As String, ColHeadName As String
Dim aCellInTbl As Boolean
Dim LastRow As Long
Dim ParseAsDataType As Byte

Application.ScreenUpdating = False

Set aCell = ActiveCell
aCellInTbl = Not aCell.ListObject Is Nothing
Set ws = ThisWorkbook.ActiveSheet
StartingPosition = ActiveCell.Address
LastRow = ws.Cells(ws.Rows.Count, aCell.Column).End(xlUp).Row

If Not IsNumeric(aCell) Then
    MsgBox "The active cell does not contain numbers. To convert between general and text formats, please select any cell with a oclumn containing numbers.", _
    vbExclamation, "Warning"
    Exit Sub
ElseIf IsEmpty(aCell) Then
    MsgBox "The active cell is blank. To convert between general and text formats, please select any cell with a oclumn containing numbers.", _
    vbExclamation, "Warning"
    Exit Sub
Else: End If

With aCell
    If .Text <> .Value Then
        ParseAsDataType = 2         'Parse as Text
    ElseIf .Text = .Value Then
        ParseAsDataType = 1         'Parse as General
    End If
End With

If aCellInTbl = True Then
    TableName = aCell.ListObject.DisplayName
    Set ActiveTable = ActiveSheet.ListObjects(TableName)
    ColHeadName = aCell.Offset(-(aCell.Row - ActiveTable.HeaderRowRange.Row), 0)
    ActiveTable.ListColumns(ColHeadName).DataBodyRange.Select
    Set ParseDataRange = Selection
Else
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Range(Selection, ActiveCell.EntireColumn).Select
    Set ParseDataRange = Selection
End If

If ParseAsDataType = 1 Then
    ParseDataRange.NumberFormat = "General"
ElseIf ParseAsDataType = 2 Then
    ParseDataRange.NumberFormat = "@"
End If

On Error GoTo 0
On Error Resume Next
With ws
    ParseDataRange.TextToColumns Destination:=ParseDataRange, FieldInfo:=Array(1, ParseAsDataType)
End With
If Err.Number = 1004 Then GoTo NoDataToParseHandler

Range(StartingPosition).Select

Set ParseDataRange = Nothing
Set ws = Nothing
Set aCell = Nothing
Set ActiveTable = Nothing

GoTo SkipErrorHandlers

'*** Error Handlers ***

NoDataToParseHandler:
Set ParseDataRange = Nothing
Set ws = Nothing
Set aCell = Nothing
Set ActiveTable = Nothing
    
MsgBox "There is no data in the selected cell(s) to parse. Please select any cell within the column you would like to parse.", _
vbCritical, "No Data to Parse"

SkipErrorHandlers:
Application.ScreenUpdating = True

End Sub