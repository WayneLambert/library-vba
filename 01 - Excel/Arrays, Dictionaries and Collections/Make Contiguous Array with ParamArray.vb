Sub PassMultipleRanges()    'Calling procedure example

Dim v As Variant
v = MakeContiguousArray(Sheet1.Range("A1:A100"), Sheet1.Range("C1:C100"))

End Sub

'Creates a contiguous array from multiple range inputs. ParamArray accepts multiple arguments
Function MakeContiguousArray(ParamArray vInput() As Variant) As Variant

Dim vOutput() As Variant
Dim i As Long, j As Long
ReDim vOutput(1 To vInput(0).Count, 0 To UBound(vInput))

For i = 0 To UBound(vInput)
    For j = 1 To vInput(i).Rows.Count
        vOutput(j, i) = vInput(i)(j)
    Next j
Next i

MakeContiguousArray = vOutput

End Function