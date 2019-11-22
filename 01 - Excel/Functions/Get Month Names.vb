Option Base 1
'Get month names from numbers populated in worksheet cells
Function GetMonthNames(Optional ByVal iMonthNo As Byte) As String

Dim AllNames As Variant
Dim iMonthVal As Byte

AllNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

If IsMissing(iMonthNo) Then
    GetMonthNames = AllNames
Else
    Select Case iMonthNo
        Case Is >= 1
        'Determine month value (for example, 13=1)
         iMonthVal = IIf(iMonthNo = 12, 12, ((iMonthNo) Mod 12))
         GetMonthNames = AllNames(iMonthVal)
      Case Is <= 0      'Vertical array
         GetMonthNames = Application.Transpose(AllNames)
     End Select
End If

End Function