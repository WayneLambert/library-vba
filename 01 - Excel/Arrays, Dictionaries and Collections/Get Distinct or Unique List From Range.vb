'This finds all of the unique entries within the Arr array. A 1D arr must be passed
'Requires a reference adding for the Microsoft Scripting Runtime library
'Removes duplicates from array using dictionary method
Function GetDistinct(tmpArr As Variant) As Variant

Dim d As Scripting.Dictionary
Set d = New Scripting.Dictionary
Dim i As Long

For i = LBound(tmpArr) To UBound(tmpArr)
    If IsMissing(tmpArr(i)) = False Then
        d.Item(tmpArr(i)) = 1
    End If
Next

GetDistinct = d.Keys

End Function

'*************************************************************************************************************************

'Creates a distinct list of items (in this case days) from a range
'The use of On Error Resume Next here is useful as ordinarily,
'adding a second item with the same key to a collection would produce an error
'On Error Resume Next enables each item within the array to be processed
'This method does not require any addtional libraries to be added to the VBA project
'This could also be written as a function to receive a range argument and return a collection

Sub GetDistinctList()

Dim collUnqDays As Collection
Dim Days As Variant
Dim i As Byte

'Read range data to array...

    'Method 1 - From an Excel Range. This creates a two dimensional array
        Days = Sheet1.Range("A2:A10").Value
    'Method 2 - From an Excel Table. This creates a two dimensional array
        'Days = Sheet1.ListObjects(1).ListColumns(1).DataBodyRange
    
    'Method 3 - From an Excel Range. The use of Application.Transpose creates a one dimensional array
        'Days = Application.Transpose(Sheet1.Range("A2:A10").Value)
    
    'Method 4 - From the Array function. This creates a one dimensional array
        'Days = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday", "Monday", "Tuesday")

Set collUnqDays = New Collection

'If using a 1D array, there is no need for the ", 1" after the array name (Days in this example)
For i = LBound(Days, 1) To UBound(Days, 1)
    On Error Resume Next
        collUnqDays.Add Item:=Days(i, 1), Key:=Days(i, 1)
    On Error GoTo 0
Next i

End Sub