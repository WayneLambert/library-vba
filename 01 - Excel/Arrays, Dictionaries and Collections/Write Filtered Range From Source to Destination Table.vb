'Calculates the column numbers to be passed through to the WriteColumnData procedure
'Where cnData ("tblData") is the source; and cnRanges ("tblAdjRanges") is the destination table
'sDIM1_HEADERNAME - sDIM4_HEADERNAME can be globally or modularly declared and assigned

Sub WriteFilteredData()

Dim lo As ListObject
Dim collColumns As Collection, v As Variant
Dim n As Byte, iColNo As Byte

Set lo = cnData.ListObjects("tblData")
Set collColumns = New Collection

With collColumns
    .Add Item:=lo.ListColumns("ID").Index, key:="ID"
    .Add Item:=lo.ListColumns(sDIM1_HEADERNAME).Index, key:=sDIM1_HEADERNAME
    .Add Item:=lo.ListColumns(sDIM2_HEADERNAME).Index, key:=sDIM2_HEADERNAME
    .Add Item:=lo.ListColumns(sDIM3_HEADERNAME).Index, key:=sDIM3_HEADERNAME
    .Add Item:=lo.ListColumns(sDIM4_HEADERNAME).Index, key:=sDIM4_HEADERNAME
End With

For Each v In collColumns
    iColNo = iColNo + 1
    Call WriteColumnData(v, iColNo)
Next v

Set collColumns = Nothing
Set lo = Nothing

End Sub

'****************************************************************************************

'Writes the filtered list for the column passed into the iColNo variable
Sub WriteColumnData(ByVal n As Byte, iColNo As Byte)

Dim lo as ListObject
Dim rFiltered As Range, rCell As Range, rStart As Range, rEnd As Range, rDest As Range
Dim ArrFilteredRange() As String

ArrFilteredRange = GetFilteredRangeArray(cnData.ListObjects("tblData").ListColumns(n). _
    DataBodyRange.SpecialCells(xlCellTypeVisible))

'Determine range size for array
Set lo = cnTRS.ListObjects("tblPropRanges")
Set rStart = lo.DataBodyRange.Cells(lo.ListRows.Count, iColNo).End(xlUp).Offset(i, 0)
Set rEnd = rStart.Offset(UBound(ArrFilteredRange) - 1, 0)
Set rDest = Range(rStart, rEnd)

'Write 1D array back to range as a vertical transposition
rDest.Value = Application.Transpose(ArrFilteredRange)

'Clean up
Erase ArrFilteredRange
Set rStart = Nothing: Set rEnd = Nothing: Set rDest = Nothing

End Sub

'****************************************************************************************

'Takes a filtered range argument and turns it into an array
Function GetFilteredRangeArray(ByRef rFilteredRange as Range) As Variant

Dim iResRow as long
Dim rArea as Range, rCell as Range
Dim Results() as Variant

ReDim Results(1 To rFilteredRange.Count)

For Each rArea in rFilteredRange.Areas
    For Each rCell in rArea.Cells
        iResRow = iResRow + 1
        Results(iResRow) = rCell.Areas(1).Value
    Next rCell
Next rArea

Set rFilteredRange = Nothing
GetFilteredRangeArray = Results

End Function