Source: https://wellsr.com/vba/2016/excel/dynamic-array-with-redim-preserve-vba/
Sub ReDimPreserve2D_AnyDimension()

Dim MyArray() As Variant
ReDim MyArray(1, 3)
'put your code to populate your array here
For i = LBound(MyArray, 1) To UBound(MyArray, 1)
    For j = LBound(MyArray, 2) To UBound(MyArray, 2)
        MyArray(i, j) = i & "," & j
    Next j
Next i

MyArray = ReDimPreserve(MyArray, 2, 4)

End Sub

'**************************************************************************************************************************************

Private Function ReDimPreserve(MyArray As Variant, nNewFirstUBound As Long, nNewLastUBound As Long) As Variant

Dim i as Long, j As Long
Dim nOldFirstUBound as Long, nOldLastUBound as Long, nOldFirstLBound as Long, nOldLastLBound As Long
Dim TempArray() As Variant 'Change this to "String" or any other data type if want it to work for arrays other than Variants.
'MsgBox UCase(TypeName(MyArray))
'---------------------------------------------------------------
'COMMENT THIS BLOCK OUT IF YOU CHANGE THE DATA TYPE OF TempArray
If InStr(1, UCase(TypeName(MyArray)), "VARIANT") = 0 Then
    MsgBox "This function only works if your array is a Variant Data Type." & vbNewLine & _
            "You have two choice:" & vbNewLine & _
            " 1) Change your array to a Variant and try again." & vbNewLine & _
            " 2) Change the DataType of TempArray to match your array and comment the top block out of the function ReDimPreserve" _
            , vbCritical, "Invalid Array Data Type"
    End
End If
'---------------------------------------------------------------
ReDimPreserve = False
'check if its in array first
If Not IsArray(MyArray) Then MsgBox "You didn't pass the function an array.", vbCritical, "No Array Detected": End

'get old lBound/uBound
nOldFirstLBound = LBound(MyArray, 1): nOldLastLBound = LBound(MyArray, 2)
nOldFirstUBound = UBound(MyArray, 1): nOldLastUBound = UBound(MyArray, 2)
'create new array
ReDim TempArray(nOldFirstLBound To nNewFirstUBound, nOldLastLBound To nNewLastUBound)
'loop through first
For i = LBound(MyArray, 1) To nNewFirstUBound
    For j = LBound(MyArray, 2) To nNewLastUBound
        'if its in range, then append to new array the same way
        If nOldFirstUBound >= i And nOldLastUBound >= j Then
            TempArray(i, j) = MyArray(i, j)
        End If
    Next
Next
'return the array redimmed
If IsArray(TempArray) Then ReDimPreserve = TempArray

End Function