Sub CallingProcForIsBetweenDates()
    Call IsBetweenDates(#1/1/2018#, #12/31/2018#, Date + 365)
End Sub

'Tests whether the dtToTest date falls between the dtStart and dtEnd
Function IsBetweenDates(dtStart As Date, dtEnd As Date, dtToTest As Date) As Boolean
    If ((dtToTest >= dtStart) And (dtToTest <= dtEnd)) Then IsBetweenDates = True
End Function