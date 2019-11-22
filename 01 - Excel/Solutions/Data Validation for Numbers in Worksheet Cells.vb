Private Sub Worksheet_Change(ByVal Target As Excel.Range)

Dim VRange As Range, cell As Range
Dim Msg As String
Dim ValidateCode As Variant

'The VRange may be set to an Excel table column (ListObject)
Set VRange = Range("InputRange")

If Intersect(VRange, Target) Is Nothing Then Exit Sub

For Each cell In Intersect(VRange, Target)
	ValidateCode = EntryIsValid(cell)
	If TypeName(ValidateCode) = "String" Then
		Msg = "Cell " & cell.Address(False, False) & ":"
		Msg = Msg & vbCrLf & vbCrLf & ValidateCode
		MsgBox Msg, vbCritical, "Invalid Entry"
		
		Application.EnableEvents = False
		cell.ClearContents
		cell.Activate
		Application.EnableEvents = True
	End If
Next cell

End Sub

'**********************************************************

Private Function EntryIsValid(cell) As Variant
'Returns True if cell is an integer between 1 and 12
'Otherwise it returns a string that describes the problem

'Tests if cell entry is a number
If Not WorksheetFunction.IsNumber(cell) Then
	EntryIsValid = "Non-numeric entry."
	Exit Function
End If

'Tests if cell entry is an integer
If CInt(cell) <> cell Then
	EntryIsValid = "Integer required."
	Exit Function
End If
		
'Tests if cell entry is between 1 and 100
If cell < 1 Or cell > 100 Then
	EntryIsValid = "Valid values are between 1 and 100."
	Exit Function
End If

'It passed all the tests
EntryIsValid = True

End Function