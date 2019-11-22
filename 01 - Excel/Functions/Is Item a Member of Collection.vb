'Checks to see if an item is a member of a collection
Private Function IsInCollection(Coln As Object, Item As String) As Boolean
    Dim Obj As Object
    On Error Resume Next
    Set Obj = Coln(Item)
    IsInCollection = Not Obj Is Nothing
End Function