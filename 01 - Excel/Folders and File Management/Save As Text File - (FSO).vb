Sub CreateTextFiles()

    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim r As Range
    Dim NoOfCols As Integer
    Dim s As String, FolderPath As String, Genre As String
    
    NoOfCols = Range("A1").CurrentRegion.Columns.Count
    
    'Change folder path as applicable
    FolderPath = Environ("UserProfile") & "\Desktop\TextFiles"
    
    If Not fso.FolderExists(FolderPath) Then fso.CreateFolder FolderPath
    
    For Each r In Range("A2", Range("A1").End(xlDown))
        
        'No of columns to offset determines the splits
        Genre = r.Offset(0, 5).Value
    
        'Output to .txt file format
        Set ts = fso.OpenTextFile(FolderPath & "\" & Genre & ".txt", ForAppending, True)
            
        s = Join(Application.Transpose(Application.Transpose(r.Resize(1, NoOfCols).Value)), vbTab)
        ts.WriteLine s

        ts.Close
    Next r
    
End Sub