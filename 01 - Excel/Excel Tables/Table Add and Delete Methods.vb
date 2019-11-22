'.Add and .Remove methods...

'Insert A New Column 4	            ActiveSheet.ListObjects("Table2").ListColumns.Add Position:=4
'Insert Column at End of Table	    ActiveSheet.ListObjects("Table2").ListColumns.Add
'Insert A New Row 5	                ActiveSheet.ListObjects("Table2").ListRows.Add (5)
'Add Row To Bottom of Table	        ActiveSheet.ListObjects("Table2").ListRows.Add AlwaysInsert:= True
'Add Totals Row	                    ActiveSheet.ListObjects("Table2").ShowTotals = True
'Remove Totals Row	                ActiveSheet.ListObjects("Table2").ShowTotals = False
'Delete Column 4	                ActiveSheet.ListObjects("Table2").ListColumns(4).Delete
'Delete Row 5	                    ActiveSheet.ListObjects("Table2").ListRows(5).Delete