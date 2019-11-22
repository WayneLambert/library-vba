'Changes the characters from start position of 3 for 10 characters within the value of the r range
'The range r must contain a hard-coded value. This will not work with a formula in there
Sub SimpleChangeOfTextInString()

Dim r As Range
Set r = ActiveCell

With r.Characters(Start:=3, Length:=10).Font
    .FontStyle = "Tahoma"
    .FontStyle = "Bold"
    .Size = 12
    .Color = vbBlue
    .ThemeFont = xlThemeFontMinor
End With

End Sub