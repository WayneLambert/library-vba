'Returns TRUE if the range name exists
Private Function RangeNameExists(sSheetName) As Boolean

Dim wbName As Name

RangeNameExists = False

For Each wbName In ThisWorkbook.Names
	If UCase$(wbName.Name) = UCase$(sSheetName) Then
		RangeNameExists = True
		Exit Function
	End If
Next wbName

End Function