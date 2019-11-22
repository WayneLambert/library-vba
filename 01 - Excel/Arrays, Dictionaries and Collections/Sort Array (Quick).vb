'Example of calling procedure...
Sub CallQuickSort()
    Dim myData() As Variant

    myData = Application.Transpose(Range("A1:A30000"))
    Call QuickSort(myData, LBound(myData), UBound(myData), False)
End Sub

'**************************************************************************************************

'Sorts a 1D array from smallest to largest using a very fast quicksort algorithm.
'Uses a conquer and divide recursive approach to sorting the array
Sub QuickSort(ByRef vArr As Variant, ByVal arrLbound As Long, ByVal arrUbound As Long, _
    Optional ByVal Ascending As Boolean = True)

Dim pvtVal As Variant, vSwap As Variant
Dim tmpLow As Long, tmpHi As Long
 
tmpLow = arrLbound: tmpHi = arrUbound
pvtVal = vArr((arrLbound + arrUbound) \ 2)
 
Do While (tmpLow <= tmpHi)                                                              '<< divide
    If Ascending = True Then
        Do While (vArr(tmpLow) < pvtVal And tmpLow < arrUbound)
           tmpLow = tmpLow + 1
        Loop
        
        Do While (pvtVal < vArr(tmpHi) And tmpHi > arrLbound)
           tmpHi = tmpHi - 1
        Loop
    Else
        Do While (vArr(tmpLow) > pvtVal And tmpLow < arrUbound)
           tmpLow = tmpLow + 1
        Loop
        
        Do While (pvtVal > vArr(tmpHi) And tmpHi > arrLbound)
           tmpHi = tmpHi - 1
        Loop
    End If
 
   If (tmpLow <= tmpHi) Then
      vSwap = vArr(tmpLow)
      vArr(tmpLow) = vArr(tmpHi)
      vArr(tmpHi) = vSwap
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
   End If
Loop
 
If (arrLbound < tmpHi) Then QuickSort vArr, arrLbound, tmpHi, Ascending                 '<< conquer
If (tmpLow < arrUbound) Then QuickSort vArr, tmpLow, arrUbound, Ascending               '<< conquer

End Sub