Sub CalculateRunTimeInSeconds()

Dim StartTime As Double, SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer

'*****************************
'Insert Code Here...
'*****************************

'Determine how many seconds code took to run
'Another EndTime variable could be used if you wanted to trap the end time
	'EndTime would replace 'Timer' in the below line
SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub

______________________

Sub CalculateRunTimeInMinutes()

Dim StartTime As Double, MinutesElapsed As String

'Remember time when macro starts
StartTime = Timer

'*****************************
'Insert Code Here...
'*****************************

'Determine how many seconds code took to run
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

'Notify user in seconds
MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub