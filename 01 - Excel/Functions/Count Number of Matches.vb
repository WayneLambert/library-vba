'In a worksheet with a codename of cnControlTotals and...
'2 ListObjects: 1) tblControlTotals and 2) tblFixedPay
'A series of fields with the same names which need to be compared to see if they match
    'In one field it is called 'Base Salary' and in the other, it contains a space before and...
    'after the field name (i.e. ' Base Salary ')

'Within the cnControlTotals worksheet
Private Sub Worksheet_Change(ByVal Target As Range)

If Not Application.Intersect(Target, Range("tblFixedPay")) Is Nothing Then
    If Target.Count > 1 Then Exit Sub
    If Target.Cells.Value = vbNullString Or IsEmpty(Target.Value) Then Exit Sub
    Call CountMatches
Else
    Exit Sub
End If

End Sub

Public Function CountMatches() As Long

Dim wb As Workbook
Dim ElementsArr() As Variant
Dim TestArr1() As Variant, TestArr2() As Variant
Dim ResultsArr(1 To 4) As Variant
Dim i As Integer, j As Integer

Set wb = Application.ThisWorkbook
ElementsArr = cnControlTotals.Range("tblControlTotalsFP[Element Desc]")

For i = LBound(ElementsArr, 1) To UBound(ElementsArr, 1)
    TestArr1 = Range("tblFixedPay[" & ElementsArr(i, 1) & "]")
    TestArr2 = Range("tblFixedPay[ " & ElementsArr(i, 1) & " ]")
    CountMatches = 0
    For j = LBound(TestArr1, 1) To UBound(TestArr1, 1)
        If TestArr1(j, 1) = vbNullString And TestArr2(j, 1) = vbNullString Then
            CountMatches = CountMatches + 1
        Else
            If Not IsError(TestArr1(j, 1)) And Not IsError(TestArr2(j, 1)) Then
                If TestArr1(j, 1) = TestArr2(j, 1) Then
                    CountMatches = CountMatches + 1
                Else
                    CountMatches = CountMatches
                End If
        Else
            CountMatches = CountMatches
            End If
        End If
    Next j
    
CountMatches = CountMatches

ResultsArr(i) = CountMatches

Next i

Set wb = Nothing

Call WriteArrayToWorkbook(ResultsArr())

End Function

Sub WriteArrayToWorkbook(ByRef ResultsArr() As Variant)

Dim Rng As Range

Set Rng = cnControlTotals.Range("tblControlTotalsFP[No Matched]")

Call InitiateMacro

Rng.Value = WorksheetFunction.Transpose(ResultsArr)
Set Rng = Nothing

Call ResumeSystemDefaults

End Sub

Sub InitiateMacro()

With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

End Sub

Sub ResumeSystemDefaults()

With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

End Sub