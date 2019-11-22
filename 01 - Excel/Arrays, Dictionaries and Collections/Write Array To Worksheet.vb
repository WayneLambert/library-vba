Option Base 1

Sub Write_ArrayToWorksheet()

Dim NoOfRows As Long
Dim NoOfCols As Long
NoOfRows = cnWriteArrayToWorksheet.Range("A1").CurrentRegion.Rows.Count - 1
NoOfCols = cnWriteArrayToWorksheet.Range("A1").CurrentRegion.Rows.Count - 1

cnWriteArrayToWorksheet.Activate

Dim c() As Variant                                  'c variable declared to contain array
ReDim c(NoOfRows, NoOfCols)
Dim RowNo As Long, ColNo As Long                    'variables to track Row and Column during loops

For RowNo = 1 To NoOfRows                           'Outer loop to work through each row
    For ColNo = 1 To NoOfCols                       'Inner loop to work through each column
        c(RowNo, ColNo) = (RowNo * ColNo)           'Formula to demonstrate times tables, for illustration
    Next ColNo
Next RowNo

cnWriteArrayToWorksheet.Range("B2").Select

For RowNo = 1 To NoOfRows
    For ColNo = 1 To NoOfCols
       ActiveCell = c(RowNo, ColNo)
       ActiveCell.Offset(0, 1).Select
    Next ColNo
    ActiveCell.Offset(1, -NoOfCols).Select
Next RowNo

End Sub

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

Sub DynamicMultiDimensionArray()

Dim CurrencyList() As Variant                                                   'Create a dynamica array with variant datatype
Dim Dimension1 As Long, Dimension2 As Long                                      'Creates variables to loop through each of the dimensions
Dim StartTime As Double, SecondsElapsed As Double                               'To record time taken

StartTime = Timer

cnCurrConv.Activate
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
End With

Dimension1 = Range("hcCurrencyCode").CurrentRegion.Rows.Count - 1
Dimension2 = Range("hcCurrencyCode").CurrentRegion.Columns.Count

ReDim CurrencyList(1 To Dimension1, 1 To Dimension2)

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

MsgBox "This dynamic multi-dimension array read and wrote successfully in " & SecondsElapsed & " seconds", vbInformation, "Sub Routine Complete"

End Sub

Sub QuickDynamicMultiDimensionArray()

Dim CurrencyList() As Variant                                                   'Create a dynamica array with variant datatype
Dim Dimension1 As Long, Dimension2 As Long                                      'Creates variables to loop through each of the dimensions
Dim StartTime As Double, SecondsElapsed As Double                               'To record time taken

StartTime = Timer

cnCurrConv.Activate
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
End With

CurrencyList = Range("hcCurrencyCode", Range("hcCurrencyCode").End(xlDown).End(xlToRight))                  'Read the array from the range within the worksheet

Range("hcCurrencyCode", Range("hcCurrencyCode").End(xlDown).End(xlToRight)).Offset(0, 5) = CurrencyList     'Wrote the array back to the worksheet at 5 columns offset to the right

Erase CurrencyList                                                              'Erases the array to release system memory

                                                                                'Test results with Timer...
With Application                                                                '1) No screenupdating = 0.06s;
    .ScreenUpdating = True                                                      '2) ScreenUpdating = 0.04s;
    .Calculation = xlCalculationAutomatic                                       '3) Calculation = 0.02s
End With

SecondsElapsed = Round(Timer - StartTime, 2)                                    'Determine how many seconds code took to run

MsgBox "This quick to code dynamic multi-dimension array read and wrote successfully in " & SecondsElapsed & " seconds.", vbInformation, "Sub Routine Complete"

End Sub

Sub CalculateWithArray()

Dim CurrencyList() As Variant, CurrencyAnswers As Variant
Dim Dimension1 As Long, Counter As Long
Dim StartTime As Double, SecondsElapsed As Double                               'To record time taken

StartTime = Timer

cnCurrConv.Activate

CurrencyList = Range("A2", Range("D2").End(xlDown))

Dimension1 = UBound(CurrencyList, 1)
Dimension2 = UBound(CurrencyList, 2)

ReDim CurrencyAnswers(1 To Dimension1, 1 To Dimension2)

For Counter = 1 To Dimension1
    CurrencyAnswers(Counter, 1) = CurrencyList(Counter, 1)
    CurrencyAnswers(Counter, 2) = CurrencyList(Counter, 2)
    CurrencyAnswers(Counter, 3) = CurrencyList(Counter, 3) * 100
    CurrencyAnswers(Counter, 4) = CurrencyList(Counter, 3) * 1000
Next Counter

Range("hcCurrencyCode", Range("hcCurrencyCode").End(xlDown).End(xlToRight)).Offset(1, 5) = CurrencyAnswers

Erase CurrencyList, CurrencyAnswers
                                                                                'Test results with Timer...
With Application                                                                '1) No screenupdating = 0.06s;
    .ScreenUpdating = True                                                      '2) ScreenUpdating = 0.04s;
    .Calculation = xlCalculationAutomatic                                       '3) Calculation = 0.02s
End With

SecondsElapsed = Round(Timer - StartTime, 2)                                    'Determine how many seconds code took to run

MsgBox "This quick to read from worksheet, calculate using array, and write back to the worksheet ran successfully in " & SecondsElapsed & " seconds.", vbInformation, "Sub Routine Complete"

End Sub