'Target Area Selection	                    Syntax
	
Dim lo As ListObject
Set lo = Sheet1.ListObjects("TableName")

lo.Range.Select                             'Entire Table	
lo.HeaderRowRange.Select	                'Header Row	    
lo.DataBodyRange.Select	                    'Data Body Range	
lo.TotalsRowRange.Select	                'Totals Row Range	
lo.ListColumns(3).Range.Select              '3rd Column	 
lo.ListRows(5).Range.Select	                '5th  Row
lo.ListColumns(3).DataBodyRange.Select      '3rd Colum (Just Data)	  