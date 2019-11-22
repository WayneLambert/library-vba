'Source:     http://allenbrowne.com/ser-36.html
'Purpose:    Return a delimited string in the date format used natively by JET SQL.
'Argument:   A date/time value.
'Note:       Returns just the date format if the argument has no time component, or a date/time format if it does.
'Author:     Allen Browne. allen@allenbrowne.com, June 2006.

Function SQLDate(vDate As Variant) As String

If IsDate(vDate) Then
    If DateValue(vDate) = vDate Then
        SQLDate = Format$(vDate, "\#mm\/dd\/yyyy\#")
    Else
        SQLDate = Format$(vDate, "\#mm\/dd\/yyyy hh\:nn\:ss\#")
    End If
End If

End Function