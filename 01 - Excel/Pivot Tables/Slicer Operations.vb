Sub SlicerOperations()
	
Dim src as SlicerCaches
Dim sls as Slicers
Dim sl as Slicer
	
'Clear pre-existing caches
For Each sc in ThisWorkbook.SlicerCaches
	sc.Delete
Next sc

'Create new slicer cache
Set scs = ThisWorkbook.SlicerCaches


'Add a slicer called 'Region' to the slicers collection
Set sls = scs.Add(cnSlicers.PivotTables("PivotTable3"), "Region", "Region").Slicers

'Add the slicer object to the interface
Set sl = sls.Add(cnSlicers, Name:="Region", Caption:="Region", _
	Top:=50, Left:=400, Width:=400, Height:=115)
	
'Deselect/Select a slicer
With scs("Region").SlicerItems("West")
	.Selected = False
	.Selected = True
End With

'Amend slicer's name
sl.Name = "Custom Name"

'Amend slicer's caption
sl.Caption = "Custom caption"

'Change style of slicer
sl.Style = "SlicerStyleLight3"

'Add/Remove the slicer's header completely
sl.DisplayHeader = False

'Adjust the row height and column width of the slicer
sl.RowHeight = 20
sl.ColumnWidth = 50

With sl.SlicerCache
	'Apply cross-filter options
	
	.CrossFilterType = xlSlicerNoCrossFilter
	.CrossFilterType = xlSlicerCrossFilterShowItemsWithNoData
	.CrossFilterType = xlSlicerCrossFilterShowItemsWithDataAtTop
	
	'Apply sort orders
	
	.SortItems = xlSlicerSortAscending
	.SortItems = xlSlicerSortDescending
	.SortUsingCustomLists = True
	.ShowAllItems = False
End With

'Reference the worksheet of the slicer
sl.Parent.Name

'Delete a slicer
sl.Delete

End Sub