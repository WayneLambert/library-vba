Option Base 1

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