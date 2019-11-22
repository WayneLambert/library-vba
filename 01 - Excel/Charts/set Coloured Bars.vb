'Sets one cell containing a formula with =REPT("|",x) with the format of 'Stencil'
'The formula must be hard coded and broken
'The RGB colours are read from a range of 6 cells with different colours in - use a grayscale
'A corresponding workbook can be found in my library through iCloud
'Can be used to show distribution
Sub SetColouredBars()

Dim r As Range
Dim sBar As String
Dim dLen As Double, iNoOfPipes As Integer, iNoOfColours As Integer, dBarLen As Long
Dim i As Integer, j As Integer: j = 1
Dim Colours As Variant

Set r = cnGraph.Range("hcBar")
r.Value = r.Value
iNoOfPipes = Len(r)

ReDim Colours(0 To 5)
For i = LBound(Colours) To UBound(Colours)
    Colours(i) = r.Offset(-3, i).Interior.Color
    dBarLen = Round(iNoOfPipes * r.Offset(-1, i).Value, 0)
    r.Characters(Start:=j, Length:=dBarLen).Font.Color = Colours(i)
    j = j + dBarLen
Next i

End Sub