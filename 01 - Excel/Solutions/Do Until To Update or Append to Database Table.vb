'In a worksheet with 4 columns storing inventory items
'Declare variables of suitable datatypes
'Use KeepSearching variable to check whether ItemID already exists within the inventory list
    'If it does, update the Stock Level with the amount that has been delivered
    'If it does not, insert a new row with the new stock item and include the product name, delivery amount and date added values
        'Increment the number of rows by 1 since there is a new stock item
'Activate the UpdateSheet
'Clear the contents from the 3 cells used to input delivery data
'Return a message box to inform the user that the inventory sheet has been updated

Sub InventoryUpdate()

Dim ItemID As Long
Dim RowNum As Long
Dim ProductName As String
Dim Delivery As Long
Dim KeepSearching As Boolean

ItemID = Worksheets("UpdateSheet").Range("A2")
RowNum = 3  'First data row in Inventory worksheet
ProductName = Worksheets("UpdateSheet").Range("B2")
Delivery = Worksheets("UpdateSheet").Range("C2")
KeepSearching = True

Worksheets("Inventory").Activate

Do Until KeepSearching = False
    If Cells(RowNum, 1).Value = ItemID Then
        Cells(RowNum, 3).Value = Cells(RowNum, 3).Value + Delivery
        Cells(RowNum, 4).Value = Date
        
        KeepSearching = False
    ElseIf Cells(RowNum, 1).Value = vbNullString Then
        Cells(RowNum, 1).Value = ItemID
        Cells(RowNum, 2).Value = ProductName
        Cells(RowNum, 3).Value = Delivery
        Cells(RowNum, 4).Value = Date
    Else
        RowNum = RowNum + 1
    End If
Loop

Worksheets("UpdateSheet").Activate
Range("A2").Select
Range("A2:C2").ClearContents

MsgBox "The inventory sheet has been updated", vbInformation, "Inventory Amended"

End Sub