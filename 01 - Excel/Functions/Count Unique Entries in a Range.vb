'Counts the unique number of entries in an input range
'Loops over each cell value (assigned to TmpArr), compares it to the previous cell.
'If different, adds one to the collection
Public Function CountUniques(rIn As Range) As Variant

Dim collUnq As New Collection
Dim rTmp As Range
Dim TmpArr As Variant, c As Variant, lc As Variant

Set rTmp = Intersect(rIn, rIn.Parent.UsedRange)
TmpArr = rTmp
On Error Resume Next
    For Each c In TmpArr
        If c <> lc Then
            If Len(CStr(c)) > 0 Then
                 collUnq.Add c, CStr(c)
            End If
        End If
        lc = c
    Next c
On Error GoTo 0

CountUniques = collUnq.Count

End Function