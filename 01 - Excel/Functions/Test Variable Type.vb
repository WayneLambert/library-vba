Sub TestVariableType()

'VarType is a reserved VBA keyword
'There are many different types including:
'vbArray, vbBoolean, vbByte, vbCurrency, vbDataObject, vbDate, vbDecimal, vbDouble, vbEmpty,
'vbError, vbInteger, vbLong, vbNull, vbSingle, vbString, vbUserDefinedType, vbVariant

Dim VariableName As Variant
	'Tests if the variable is a boolean type and exits the sub routine if it is
	'Of course, vbBoolean could be replaced with any of the other types
	If VarType(VariableName) = vbBoolean then Exit Sub
End Sub

'Alternatively, the above procedure could be written as a function...

Function IsBoolean(ByRef VariableName as Variant) as Boolean

If VarType(VariableName) = vbBoolean then
	IsBoolean = True
Else
	IsBoolean = False
End If

End Function