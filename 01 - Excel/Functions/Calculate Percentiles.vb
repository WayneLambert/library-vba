Function GetPercentileValue() As Double
    Dim TestArr() As Variant
    Dim dPercVal As Double
    Dim dResult As Double

    'Percentile to calculate - i.e. the kth value
    dPercVal = 0.75
    'Read range values into dynamic array
    TestArr() = cnControlTotals.Range("tblFixedPay[Base Salary]")
    'Calculate percentile value
    dResult = Application.WorksheetFunction.Percentile((TestArr), dPercVal)
End Function

'Calculates the percentile using the Application.WorksheetFunction function
Function CalculatePercentile()
    Dim dPercVal As Double

    dPercVal = Application.WorksheetFunction.Percentile(Range("tblFixedPay[Base Salary]"), 0.5)
    CalculatePercentile = dPercVal
End Function

'Calculates the percentile using the Evaluate function
Function CalculatePercentile()
    Dim dPercVal As Double

    dPercVal = Evaluate("Percentile(tblFixedPay[Base Salary], 0.5)")
    MyPercentile = dPercVal
End Function