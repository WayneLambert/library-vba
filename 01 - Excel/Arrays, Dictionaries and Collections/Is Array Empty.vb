Call IsArrayEmpty(Arr:=Arr)

Function IsArrayEmpty(ByVal Arr as Variant) As Boolean)
    IsArrayEmpty = Len(Join(Arr,""))=0
End Function