'Transposing an array changes its orientation and makes it one dimensional
Sub TransposeArray()

Dim loMyTable as ListObject
Dim ArrHeaders as Variant

ArrHeaders = Application.Transpose(Application.Transpose(loMyTable.HeaderRowRange))

End Sub