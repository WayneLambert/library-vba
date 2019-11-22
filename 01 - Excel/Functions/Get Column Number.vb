'Gets the column letter according to the long number passed into the function
Function GetColumnLetter(ByVal ColNum As Long) As String
    GetColumnLetter = Split(Cells(1, ColNum).Address(True, False), "$")(0)
End Function