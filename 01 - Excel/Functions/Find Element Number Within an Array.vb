'Example sub procedure that reads a range into a 1D array
Sub FindElementNumber()

Dim TmpArr As Variant
Dim iElementPosition As Long

TmpArr = Application.Transpose(Sheet1.Range("A2:A30000"))
iElementPosition = GetElementNumberInArray(TmpArr, 9006404)

End Sub

'*************************************************************************************

'Function to find the element number where a value is located. 1D array required
Function GetElementNumberInArray(ByRef arr As Variant, ByVal vFind As Variant) As Long

Dim i As Long

For i = LBound(arr) To UBound(arr)
    If arr(i) = vFind Then
        GetElementNumberInArray = i
        Exit Function
    End If
Next i

GetElementNumberInArray = Null

End Function