Sub CreatePivotTable()

Dim PTcache As PivotCache
Dim pt As PivotTable

Application.ScreenUpdating = False
'Delete PivotSheet if it exists
On Error Resume Next
Application.DisplayAlerts = False
Sheets("Report").Delete
On Error GoTo 0

'Create a Pivot Cache
Set PTcache = ActiveWorkbook.PivotCaches.Create( _
  SourceType:=xlDatabase, _
  SourceData:=Range("A1").CurrentRegion.Address)

'Add new worksheet
Worksheets.Add After:=ThisWorkbook.ActiveSheet
ActiveSheet.Name = "Report"
ActiveWindow.DisplayGridlines = False
ActiveWindow.DisplayHeadings = False

'Create the Pivot Table from the Cache
Set pt = ActiveSheet.PivotTables.Add( _
  PivotCache:=PTcache, _
  TableDestination:=Range("B2"), _
  TableName:="BudgetPivot")

With pt
'Add fields
	.PivotFields("Category").Orientation = xlPageField
	.PivotFields("Division").Orientation = xlPageField
	.PivotFields("Department").Orientation = xlRowField
	.PivotFields("Month").Orientation = xlColumnField
	.PivotFields("Budget").Orientation = xlDataField
	.PivotFields("Actual").Orientation = xlDataField
	.DataPivotField.Orientation = xlRowField
'Add a calculated field to compute variance
	.CalculatedFields.Add Name:="Variance", Formula:="=Budget-Actual"
	.PivotFields("Variance").Orientation = xlDataField
'Specify a number format
	.DataBodyRange.NumberFormat = "#,##0.00;[Red](#,##0.00)"
'Apply a style
	.TableStyle2 = "PivotStyleMedium2"
	.MergeLabels = True
'Hide Field Headers
	.DisplayFieldCaptions = False
'Change the captions
	.PivotFields("Sum of Budget").Caption = " Budget"
	.PivotFields("Sum of Actual").Caption = " Actual"
	.PivotFields("Sum of Variance").Caption = " Variance"
End With

'Add a little whitespace in the top left corner of the worksheet to improve readability
Range("1:1").Rows.Insert
Range("A:A").ColumnWidth = 1
Range("1:1").RowHeight = 5

End Sub