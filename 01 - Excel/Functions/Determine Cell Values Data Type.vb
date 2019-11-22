'Returns the cell type of the upper left cell in a range
Function CellType(ByRef Rng as Range) As String

Dim r As Range
Set r = Rng.Range("A1")

Select Case True
	Case IsEmpty(r)
	 CellType = "Blank"
	Case r.NumberFormat = "@"
	 CellType = "Text"
	Case Application.IsText(r)
	 CellType = "Text"
	Case Application.IsLogical(r)
	 CellType = "Logical"
	Case Application.IsErr(r)
	 CellType = "Error"
	Case IsDate(r)
	 CellType = "Date"
	Case InStr(1, r.Text, ":") <> 0
	 CellType = "Time"
	Case IsNumeric(r)
	 CellType = "Number"
End Select

End Function