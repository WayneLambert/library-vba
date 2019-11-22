' Loops through a specified column, and cuts each set of distinct values into a separate workbook by creating a copy and deleting rows below and above
' Values in the selected column should be sorted or unique
' The following cells are ignored when delimiting sections:
' - blank cells, or containing spaces only
' - same value repeated
' - cells containing "total"
' Files are saved in a "File Splits" subfolder from the location of the source workbook, and named after the section name.
Public Sub CutFiles()

Dim oWS As Worksheet ' Original sheet
Dim RowNo As Long, ColNo As Long, FirstRowNo As Long, StartRowNo As Long, StopRowNo As Long, TotalRowsNo As Long, NoOfFiles As Long
Dim strSectionName As String, strCell As String, strFilePath As String
Dim rngCell As Range
Dim oWB As Workbook

ColNo = Application.InputBox("Enter the column number you would like to split the files with.", "Select column...", 2, , , , , 1)
RowNo = Application.InputBox("Enter the first data row.", "Select row...", 5, , , , , 1)
FirstRowNo = RowNo

Set oWS = Application.ActiveSheet
Set oWB = Application.ActiveWorkbook
TotalRowsNo = oWS.UsedRange.Rows.Count
strFilePath = Application.ActiveWorkbook.Path

If Dir(strFilePath + "\File Splits", vbDirectory) = "" Then
    MkDir strFilePath + "\File Splits"
End If

With Application
    .EnableEvents = False
    .ScreenUpdating = False
End With

Do
    Set rngCell = oWS.Cells(RowNo, ColNo)
    strCell = Replace(rngCell.Text, " ", "")

    If strCell = "" Or (rngCell.Text = strSectionName And StartRowNo <> 0) Or InStr(1, rngCell.Text, "total", vbTextCompare) <> 0 Then
        ' No actions are required when condition is met
    Else
        If StartRowNo = 0 Then                  ' Found new section
            strSectionName = rngCell.Text       ' StartRow delimiter not set, meaning beginning a new section
            StartRowNo = RowNo
        Else
            StopRowNo = RowNo - 1               ' StartRow delimiter set, meaning we reached the end of a section
            CopySheet oWS, FirstRowNo, StartRowNo, StopRowNo, TotalRowsNo, strFilePath, strSectionName, oWB.fileFormat      ' Pass variables to a separate sub to create and save the new worksheet
            NoOfFiles = NoOfFiles + 1           ' Reset section delimiters
            StartRowNo = 0
            StopRowNo = 0                       ' Ready to continue loop
            RowNo = RowNo - 1
        End If
    End If

    If RowNo < TotalRowsNo Then                 ' Continue until last row is reached
            RowNo = RowNo + 1
    Else                                        ' Finished. Save the last section
        StopRowNo = RowNo
        CopySheet oWS, FirstRowNo, StartRowNo, StopRowNo, TotalRowsNo, strFilePath, strSectionName, oWB.fileFormat
        NoOfFiles = NoOfFiles + 1
        Exit Do                                 ' Exit
    End If
Loop

With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With

MsgBox Str(NoOfFiles) & " files have been cut and saved at " & strFilePath, vbInformation, "Process Complete"

End Sub

Public Sub DeleteRows(aWS As Worksheet, FromRowNo As Long, ToRowNo As Long)

Dim Rng As Range

Set Rng = Range(aWS.Cells(FromRowNo, 1), aWS.Cells(ToRowNo, 1)).EntireRow

With Rng
    .Select
    .Delete
End With

End Sub

Public Sub CopySheet(oWS As Worksheet, FirstRowNo As Long, StartRowNo As Long, StopRowNo As Long, TotalRowsNo As Long, strFilePath As String, strSectionName As String, fileFormat As XlFileFormat)

Dim aWS As Worksheet ' Copied sheet
Dim aWB As Workbook ' New workbook

oWS.Copy                                ' Copy worksheet
Set aWS = Application.ActiveSheet

If TotalRowsNo > StopRowNo Then         ' Delete Rows after section
    Call DeleteRows(aWS, StopRowNo + 1, TotalRowsNo)
End If

If StartRowNo > FirstRowNo Then         ' Delete Rows before section
    Call DeleteRows(aWS, FirstRowNo, StartRowNo - 1)
End If

aWS.Cells(1, 1).Select          ' Select left-topmost cell

' Clean up a few characters to prevent invalid filename
strSectionName = Replace(strSectionName, "/", " ")
strSectionName = Replace(strSectionName, "\", " ")
strSectionName = Replace(strSectionName, ":", " ")
strSectionName = Replace(strSectionName, "=", " ")
strSectionName = Replace(strSectionName, "*", " ")
strSectionName = Replace(strSectionName, ".", " ")
strSectionName = Replace(strSectionName, "?", " ")

aWS.SaveAs strFilePath + "\File Splits\" + strSectionName, fileFormat           ' Save in same format as original workbook

Set aWB = aWS.Parent
aWB.Close SaveChanges:=False                                                    ' Close without saving changes

End Sub