Public Class ListInventory_Data
    Public InventoryID As Integer
    Public InventoryCode As String
    Public InventoryName As String

    Public Shared Function NewListInventory(ByVal invID As Integer, ByVal invCode As String, ByVal invName As String) As ListInventory_Data
        Dim iData As New ListInventory_Data
        iData.InventoryID = invID
        iData.InventoryCode = invCode
        iData.InventoryName = invName
        Return iData
    End Function
End Class