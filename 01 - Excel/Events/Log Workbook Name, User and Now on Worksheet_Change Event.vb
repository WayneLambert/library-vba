'Whenever a change occurs in column B, sub runs to completion, worksheet change event exits at start of procedure...
'if worksheet_change event does not relate to column B
Private Sub Worksheet_Change(ByVal Target As Range)

Dim TargetRow As Integer

TargetRow = Target.Row

If Target.Column = 2 Then Exit Sub

If Target.Column = 2 Then
	If Target.Value <> "" Then
		Range("M" & TargetRow) = Application.ThisWorkbook.Name
		Range("N" & TargetRow) = Application.UserName
		Range("O" & TargetRow) = Now()
	Else: End If
End If

If Target.Column = 2 Then
	If Target.Value = "" Then
		Range("M" & TargetRow) = ""
		Range("N" & TargetRow) = ""
		Range("O" & TargetRow) = ""
	Else: End If
End If