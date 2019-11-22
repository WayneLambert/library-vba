'To delete all rows from a spreadsheet using VBA,
'work from the bottom of the range to the top through each iteration
'with a For Loop with a Step of -1 to decrement
'Ref: Boris Paskhaver, Excel VBA Programming - The Complete Guide, Udemy. Lecture 99

Sub DeleteRows()

Dim iLastRow as Long, i as Long

iLastRow = Cells(rows.count,1).End(xlUp).Row

For i = iLastRow to 1 Step -1
    If Cells(i,1)="DELETE" Then
        Cells(i,1).EntireRow.Delete    
    End If
Next i

End Sub