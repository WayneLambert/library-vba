'Checks to see if formulas are still calculating before resuming
Sub WaitUntilFinished()

Application.Calculate 	'Optional - recalculates all formulas
If Not Application.CalculationState = xlDone Then DoEvents
'Insert remaining code here

End Sub

'****************************************************************

'An alternative on the above sub routine is to use a loop
Sub WaitUntilFinishedLoop()

'Loop until all your calculations are done
Application.Calculate 	'Optional - recalculates all formulas
Do Until Application.CalculationState = xlDone
    DoEvents
Loop
'Insert remaining code here

End Sub