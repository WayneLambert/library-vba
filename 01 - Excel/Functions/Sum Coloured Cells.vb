'Source: Leila Gharani - Unlock Excel VBA and Excel Macros, Udemy Course: Lecture 160
'Calculates the sum of numbers within a range passed to the function that have the same
'colour as another range that is passed into the function
'If cell colour is determined by a conditional format, use r.DisplayFormat.Interior instead
Function SumColour(ByRef RangeWithColour As Range, ByRef RangeToSum as Range) As Double

Dim r as Range
Dim ColourToMatch as Long

ColourToMatch = RangeWithColour.Cells(1,2,).Interior.Color

For Each r in RangeToSum
    If r.Interior.Color = ColourToMatch Then 
        SumColour = SumColour + r.Value
    End If
Next r

End Function

'***********************************************************************

'To activate this when changes are made, it needs to be triggered when a
'Worksheet_SelectionChange event fires
Private Sub Worksheet_SelectionChange(ByVal Target as Range)
    ActiveSheet.Calculate
End Sub