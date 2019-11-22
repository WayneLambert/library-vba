'Speeds up code execution
Public Sub FastWB(Optional ByVal Opt As Boolean = True)

With Application
    .Calculation = IIf(Opt, xlCalculationManual, xlCalculationAutomatic)
    If .DisplayAlerts <> Not Opt Then .DisplayAlerts = Not Opt
    If .DisplayStatusBar <> Not Opt Then .DisplayStatusBar = Not Opt
    If .EnableAnimations <> Not Opt Then .EnableAnimations = Not Opt
    If .EnableEvents <> Not Opt Then .EnableEvents = Not Opt
    If .ScreenUpdating <> Not Opt Then .ScreenUpdating = Not Opt
End With
    'If the ws argumenet is missing, it will turn all features on and off
    'for all WorkSheets in the collection
    Call FastWS(, Opt:=True)

End Sub

Public Sub FastWS(Optional ByVal ws As Worksheet, Optional ByVal Opt As Boolean = True)

If ws Is Nothing Then
    For Each ws In Application.ThisWorkbook.Sheets
        Call OptimiseWS(ws, Opt:=True)
    Next ws
Else
    Call OptimiseWS(ws, Opt:=True)
End If

End Sub

Private Sub OptimiseWS(ByVal ws As Worksheet, ByVal Opt As Boolean)

With ws
    .DisplayPageBreaks = False
    .EnableCalculation = Not Opt
    .EnableFormatConditionsCalculation = Not Opt
    .EnablePivotTable = Not Opt
End With

End Sub