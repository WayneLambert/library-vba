Sub SetPageBreaks()

wsRows As Long, pbRows As Long
 
'Set the number of rows to create a new page break for printing
pbRows = 30

'Optional area to set the page up so it is printer ready
With ActiveSheet.PageSetup
    .CenterHorizontally = False
    .CenterVertically = False
    .Orientation = xlLandscape
    .FitToPagesWide = 1
    .FitToPagesTall = False
End With

'For every 30 rows after the 31st row, a page break is inserted
'The starting point for wsRows can be adjusted if the first page is different. For example, using a report title on first page  
For wsRows = 31 To 1026 Step pbRows     '1026 is the maximum number of page breaks in a worksheet permissible by Excel
    ActiveSheet.HPageBreaks.Add Before:=ActiveSheet.Cells(wsRows, 1)
Next
    
End Sub