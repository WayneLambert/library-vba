Sub ConvertMinutesToHours()

Dim Mins() as variant
Dim i as Integer

Mins = Range("A2",Range("A1").End(xlDown))

For i = LBound(Mins, 1) to UBound(Mins, 1)
	Mins(i,1) = MinsToHours(Mins(i, 1))
Next i

'Passes array values stored in memory
'Determines size of range to be used by resizing it based upon the
	'arrays upper bounds
Range("B2").Resize = UBound(Mins, 1),1) = Mins

End Sub

'Converts minutes to hours and minutes
'Uses byVal to pass argument to function correctly
Function MinsToHours (ByVal m as Integer) as String

	MinsToHours = Int(m / 60) & "h " & (m Mod 60) & "m"
	
End Function