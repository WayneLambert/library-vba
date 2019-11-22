Sub ConsolidateQuestionnaireResults()

Dim wbToCollate As Workbook, wbMaster As Workbook, Rng As Range, SortCell As Range, SortTable As Range
Dim wbToCollatePath As String, FileToProcess As String, FileExt As String, rComments As String, PasteAddress As String, SeqNo As String
Dim ColNo As Integer, PasteRow As Integer, NoCompleted As Integer

With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlManual
End With

Set wbMaster = Application.ThisWorkbook
wbToCollatePath = wbMaster.Worksheets("Results").Range("hcFilepath")
wbToCollatePath = IIf(Right(wbToCollatePath, 1) <> "\", wbToCollatePath & "\", wbToCollatePath)
If PathExists(wbToCollate) = False Then GoTo PathDoesNotExistHandler
FileExt = "*.xls*"          'Target file extension (must include wildcard "*")
FileToProcess = Dir(wbToCollate & FileExt)  'Target pathway with ending extension
PasteRow = 3
NoCompleted = 0
PasteAddress = "C" & PasteRow

'Loop through each Excel file in folder
Do While FileToProcess <> ""
    Set wbToCollate = Workbooks.Open(Filename:=wbToCollatePath & FileToProcess)
    SeqNo = Left(FileToProcess, 3)
    DoEvents        'Ensure Workbook has opened before moving on to next line of code
    With wbToCollate.Worksheets("Checklist")
        Set Rng = .Range("D3:E24")
        rComments = .Range("C26").Value
    End With
    With wbMaster.Worksheets("Results").Range(PasteAddress)
        .Resize(Rng.Columns.Count, Rng.Rows.Count).Cells.Value = Application.Transpose(Rng.Cells.Value)
        .Offset(0, -2).Value = SeqNo
        .Offset(0, -1).Value = "Manager"
        .Offset(1, -2).Value = SeqNo
        .Offset(1, -1).Value = "Employee"
        .Offset(0, 23).Value = rComments
    End With
    wbToCollate.Close SaveChanges:=False
    PasteRow = PasteRow + 2
    PasteAddress = "C" & PasteRow
    NoCompleted = NoCompleted + 1
    DoEvents        'Ensure the workbook has closed before moving onto the next line
    FileToProcess = Dir     'Get next filename
Loop

Set Rng = Nothing
Set wbToCollate = Nothing

With wbMaster.Worksheets("Results").Range("hcFilepath")
    .Select
    .ClearContents
    .Interior.Color = RGB(255, 255, 255)
End With

Set wbMaster = Nothing

Dim OneRange As Range
Dim aCell As Range

Set SortTable = Range("tblResults")
Set SortCell = Range("A2")

SortTable.Sort Key1:=SortCell, Order1:=xlAscending, Header:=xlYes
SortCell.Select

With Application
    .EnableEvents = True
    .Calculation = xlAutomatic
    .ScreenUpdating = True
End With

MsgBox NoCompleted & "files have been consolidated.", vbInformation, "Consolidation Complete"
Exit Sub

PathDoesNotExistHandler:

Set Rng = Nothing
Set wbToCollate = Nothing
With wbMaster.Worksheets("Results").Range("hcFilepath")
    .Select
    .Interior.Color = RGB(255, 0, 0)
End With
Set wbMaster = Nothing
With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlManual
End With

MsgBox "This path does not exist.", vbCritical, "File Path Error"
Exit Sub

End Sub

Private Function PathExists(wbToCollatePath) As Boolean

If Dir(wbToCollatePath) = "" Then
    PathExists = False
Else
    PathExists = (GetAttr(wbToCollatePath) And vbDirectory) = vbDirectory
End If

End Function