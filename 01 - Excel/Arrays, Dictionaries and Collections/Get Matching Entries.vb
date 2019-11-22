'Can be used to get the matching hierarchy by creating a list of all of the values beneath that match the sTest for the hierarchy level above
'Can be used for cascading drop down selections
Function GetMatchingEntries(ByVal Arr As Variant, ByVal sTest As String) As Variant

Dim dictUnq As Scripting.Dictionary, key As Variant
Dim i As Long

Set dictUnq = New Scripting.Dictionary

For i = LBound(Arr) To UBound(Arr)
    If Arr(i, 1) = sTest Then dictUnq(Arr(i, 2)) = 1
Next i

Erase Arr

'Can be used in conjunction with the SortArray function to get the array in ascending alphabetical order
GetMatchingEntries = SortArray(ArrayToSort:=dictUnq.Keys, bAscending:=True)

End Function