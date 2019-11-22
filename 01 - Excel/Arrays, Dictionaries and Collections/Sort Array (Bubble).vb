'Call the function as array = SortArray function and pass through either an array, the keys of a dictionary or collection
'Example of call below. For arrays of more than 5,000 elements, use the QuickSort approach instead
CutOptions = SortArray(ArrayToSort:=dictCutOptions.Keys, Ascending:=True)

Public Function SortArray(ArrayToSort, Ascending As Boolean)

Dim sTmp As Variant
Dim i As Long, j As Long

If Ascending = True Then
    For i = LBound(ArrayToSort) To UBound(ArrayToSort)
         For j = i + 1 To UBound(ArrayToSort)
             If ArrayToSort(i) > ArrayToSort(j) Then
                 sTmp = ArrayToSort(j)
                 ArrayToSort(j) = ArrayToSort(i)
                 ArrayToSort(i) = sTmp
             End If
         Next j
     Next i
Else
    For i = LBound(ArrayToSort) To UBound(ArrayToSort)
         For j = i + 1 To UBound(ArrayToSort)
             If ArrayToSort(i) < ArrayToSort(j) Then
                 sTmp = ArrayToSort(j)
                 ArrayToSort(j) = ArrayToSort(i)
                 ArrayToSort(i) = sTmp
             End If
         Next j
     Next i
End If

SortArray = ArrayToSort

End Function