Option Base 1

Sub MultiDimensionArray()

Dim CurrencyList(1 To 166, 1 To 4)                                              'Create two dimensional array with fixed dimensions
Dim Dimension1 As Long, Dimension2 As Long                                      'Creates variables to loop through each of the dimensions
Dim StartTime As Double, SecondsElapsed As Double                               'To record time taken

StartTime = Timer

cnCurrConv.Activate
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
End With

For Dimension1 = LBound(CurrencyList, 1) To UBound(CurrencyList, 1)             'Defines the looping range for the first dimension from its lower to upper bounds
    For Dimension2 = LBound(CurrencyList, 2) To UBound(CurrencyList, 2)         'Defines the looping range for the second dimension from its lower to upper bounds
        CurrencyList(Dimension1, Dimension2) = _
        Range("hcCurrencyCode").Offset(Dimension1, Dimension2 - 1)              'The CurrencyList array reads information ... from worksheet's corresponding slot
    Next Dimension2                                                             'Moves second dimension loop onto its next iteration
Next Dimension1                                                                 'Moves first dimension loop onto its next iteration

cnCurrConv.Range("hcCurrencyCode").Offset(1, 4).Activate                        'Offsets 4 positions to the right of the range A2

For Dimension1 = LBound(CurrencyList, 1) To UBound(CurrencyList, 1)             'Defines the looping range for the first dimension from its lower to upper bounds
    For Dimension2 = LBound(CurrencyList, 2) To UBound(CurrencyList, 2)         'Defines the looping range for the second dimension from its lower to upper bounds
         ActiveCell.Offset(Dimension1 - 1, Dimension2) = _
         CurrencyList(Dimension1, Dimension2)                                   'The CurrencyList array writes it back ... to the corresponding activecell on the worksheet
    Next Dimension2                                                             'Moves second dimension loop onto its next iteration
Next Dimension1                                                                 'Moves first dimension loop onto its next iteration

Erase CurrencyList                                                              'Erases the array to release system memory

                                                                                'Test results with Timer...
With Application                                                                '1) No screenupdating = 0.06s;
    .ScreenUpdating = True                                                      '2) ScreenUpdating = 0.04s;
    .Calculation = xlCalculationAutomatic                                       '3) Calculation = 0.02s
End With

SecondsElapsed = Round(Timer - StartTime, 2)                                    'Determine how many seconds code took to run

MsgBox "This fixed multi-dimension array read and wrote successfully in " & SecondsElapsed & " seconds", vbInformation, "Sub Routine Complete"

End Sub