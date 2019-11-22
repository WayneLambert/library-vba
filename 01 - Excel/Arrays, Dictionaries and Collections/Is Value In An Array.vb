'Checks the see if a value is in an array of values
Private Function IsInArray(ByVal valToFind as Variant, ByRef ArrToCheck as Variant) as Boolean

Dim e as variant

On Error GoTo IsInArrayError		'Array is empty
	For Each e in ArrToCheck
		If e = valToFind Then
			IsInArray=True
			Exit Function
		End If
	Next e
IsInArrayError:
	On Error GoTo 0
	IsInArray = False
End Function