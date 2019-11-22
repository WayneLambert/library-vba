'This procedure creates 2 pivot tables for every Likert scale question
'within the ActiveRegion of the SurveyData worksheet
Sub MakeSurveyResultsPivotTables()

Dim PTCache As PivotCache
Dim pt As PivotTable
Dim SummarySheet As Worksheet
Dim ItemName As String
Dim Row As Long, Col As Long, i As Long, wsRows As Long, pbRows As Long, NoOfQuestions As Long
 
Application.ScreenUpdating = False
 
'Delete Summary sheet if it exists
 On Error Resume Next
 Application.DisplayAlerts = False
 Sheets("Survey Summaries").Delete
 On Error GoTo 0
 
'Add Summary sheet
 Set SummarySheet = Worksheets.Add(After:=ThisWorkbook.ActiveSheet)
 SummarySheet.Name = "Survey Summaries"
 
'Create Pivot Cache
 Set PTCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
 SourceData:=Sheets("SurveyData").Range("A1").CurrentRegion)
 
'Sets row initially to 1 to place first title of pivot table for first question
'This is incremented by 10 each time to represent the number of rows required for a block
Row = 1
'Sets the number of questions within the SurveyData ActiveRegion data range
NoOfQuestions = ThisWorkbook.Sheets("SurveyData").Range("A1").CurrentRegion.Columns.Count - 2
'Set the number of rows to create a new page break for printing
pbRows = 30

For i = 1 To NoOfQuestions
    For Col = 1 To 6 Step 5 '2 columns
        ItemName = Sheets("SurveyData").Cells(1, i + 2)
        With Cells(Row, 1)  'Use Col instead of 1 if preference is to display title above both tables
            .Value = ItemName
            .Font.Size = 16
        End With
        
        'Create pivot table
        Set pt = ActiveSheet.PivotTables.Add( _
        PivotCache:=PTCache, _
        TableDestination:=SummarySheet.Cells(Row + 1, Col))
        
        'Add the fields
        If Col = 1 Then 'Frequency tables
            With pt.PivotFields(ItemName)
                .Orientation = xlDataField
                .Function = xlCount
                .Name = "Frequency"
            End With
            Else ' Percent tables
                With pt.PivotFields(ItemName)
                .Orientation = xlDataField
                .Function = xlCount
                .Name = "Percent"
                .Calculation = xlPercentOfColumn
                .NumberFormat = "0.0%"
            End With
        End If
        
        With pt
            .PivotFields(ItemName).Orientation = xlRowField
            .PivotFields("Sex").Orientation = xlColumnField
            .PivotFields("Sex").AutoSort xlDescending, "Sex"
            .TableStyle2 = "PivotStyleMedium2"
            .DisplayFieldCaptions = False
        End With
        
        If Col = 6 Then
            'Add data bars to the last column
            With pt
                .ColumnGrand = False
                .DataBodyRange.Columns(3).FormatConditions.AddDatabar
                With pt.DataBodyRange.Columns(3).FormatConditions(1)
                    .BarFillType = xlDataBarFillSolid
                    .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
                    .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
                End With
            End With
        End If
    Next Col
    Row = Row + 10
Next i

'   Replace numbers with descriptive text
With Range("A:A,F:F")
    .Replace What:=1, Replacement:="Strongly Agree", LookAt:=xlWhole
    .Replace What:=2, Replacement:="Agree", LookAt:=xlWhole
    .Replace What:=3, Replacement:="Neutral", LookAt:=xlWhole
    .Replace What:=4, Replacement:="Disagree", LookAt:=xlWhole
    .Replace What:=5, Replacement:="Strongly Disagree", LookAt:=xlWhole
End With

'Set screen to be visually appealing
With ActiveWindow
    .DisplayGridlines = False
    .DisplayHeadings = False
End With

'Add a little whitespace in the top left corner of the worksheet to improve readability
Range("A:A").Columns.Insert (1)
Range("A:A").ColumnWidth = 1
Range("B:B,G:G").ColumnWidth = 15
Range("C:D,H:I").ColumnWidth = 6.5
Range("F:F").ColumnWidth = 1

'Sets the page up so it is printer ready
With SummarySheet.PageSetup
    .CenterHorizontally = False
    .CenterVertically = False
    .Orientation = xlLandscape
    .FitToPagesWide = 1
    .FitToPagesTall = False
End With
   
For wsRows = 31 To 1026 Step pbRows     '1026 is the maximum number of page breaks in a worksheet permissible by Excel
    SummarySheet.HPageBreaks.Add Before:=SummarySheet.Cells(wsRows, 1)
Next
    
End Sub