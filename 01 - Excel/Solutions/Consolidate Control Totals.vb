'This sub was used to consolidate control totals in the actuals v proposals process
Sub ConsolidateControlTotals()

Const sBASE_PATH As String = "E:\YE 2017\AvP Process\Country Folders\"
Const sFILE_NAME As String = "01. Year End Payment Checks Return Files - "
Const iFP_ROW_OFFSET As Integer = 23
Dim wbCountry As Workbook
Dim ws As Worksheet
Dim i As Single
Dim sPasteRangeAdd As String
Dim Countries() As Variant
Dim loCopyTable As ListObject
Dim dStartTime As Double, dEndTime As Double, dTimeTaken As Double, sTimeTaken As String

'Start timer
dStartTimer = Timer

'Amend application to speed up code execution
'Call StartExecution

'Read country data range into 1 dimensional array
Countries = Application.Transpose(cnRefTables.Range("tblCountries[Host Country]"))

'Loop over each country in the array
For i = LBound(Countries) To UBound(Countries)
    Set wbCountry = Workbooks.Open(Filename:=sBASE_PATH & Countries(i) & "\" & _
        sFILE_NAME & Countries(i) & ".xlsb", ReadOnly:=True)
    'Ensure the file is open before resuming code execution
    DoEvents
    'Set the worksheet to the "Control Totals" worksheet for the country being processed in the loop (i.e. country of i)
    Set ws = wbCountry.Worksheets("Control Totals")
    ' (1) For Fixed Pay
        Set loCopyTable = ws.ListObjects("tblControlTotalsFP")
        sPasteRangeAdd = "B" & (i * iFP_ROW_OFFSET) - iFP_ROW_OFFSET + 2 & ":K" & (i * iFP_ROW_OFFSET) + 1
        With cnConsFP
            .Range(sPasteRangeAdd).Value = loCopyTable.DataBodyRange.Value
            sPasteRangeAdd = "A" & (i + iFP_ROW_OFFSET) - iFP_ROW_OFFSET + 2 & ":A" & (i * iFP_ROW_OFFSET) + 1
            .Range(sPasteRangeAdd).Value = Countries(i)
        End With
    ' (2) For VC
        Set loCopyTable = ws.ListObjects("tblControlTotalsCash")
        sPasteRangeAdd = "B" & i + 1 & ":K" & i + 1
        With cnConsVC
            .Range(sPasteRangeAdd).Value = loCopyTable.DataBodyRange.Value
            .Range("A1").Offset(i, 0).Value = Countries(i)
        End With
    ' (1) For Recognition
        Set loCopyTable = ws.ListObjects("tblControlTotalsRec")
        sPasteRangeAdd = "B" & i + 1 & ":K" & i + 1
        With cnConsRec
            .Range(sPasteRangeAdd).Value = loCopyTable.DataBodyRange.Value
            .Range("A1").Offset(i, 0).Value = Countries(i)
        End With
    Set ws = Nothing
    wbCountry.Close SaveChanges:=False
Next i

'End the timer and set the string in mm:ss format
dEndTime = Timer
dTimeTaken = Int(dEndTime - dStartTime)
sTimeTaken = Format(dTimeTaken / 60, "00") & ":" & Format(dTimeTaken Mod 6, "00")

MsgBox "Consolidation has ran for " & i & " files in " & sTimeTaken & ".", vbInformation, "Process Complete"

'Returns the application to its default configuration
'Call ReturnToDefaults

End Sub