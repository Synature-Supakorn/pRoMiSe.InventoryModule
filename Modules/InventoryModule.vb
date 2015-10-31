Module InventoryModule

    Friend Function ListInventoryName(ByVal globalVariable As GlobalVariable, ByRef invList As List(Of ListInventory_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtResult As DataTable
        invList = New List(Of ListInventory_Data)
        Try
            dtResult = InventorySQL.ListInventory(globalVariable.DocDBUtil, globalVariable.DocConn, globalVariable.ORDERINVENTORY_BYID)
            invList = New List(Of ListInventory_Data)
            For i = 0 To dtResult.Rows.Count - 1
                invList.Add(ListInventory_Data.NewListInventory(dtResult.Rows(i)("ShopID"), dtResult.Rows(i)("ShopCode"), dtResult.Rows(i)("ShopName")))
            Next i
        Catch ex As Exception
            resultText = ex.ToString
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "InventoryModule", "ListInventoryName", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function ListInventoryViewForSelectInventory(ByVal globalVariable As GlobalVariable, ByVal viewFromInventoryID As Integer, ByRef invList As List(Of ListInventory_Data),
                                                           ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtResult As DataTable
        invList = New List(Of ListInventory_Data)
        Try
            dtResult = InventorySQL.ListInventoryViewForSelectShop(globalVariable.DocDBUtil, globalVariable.DocConn, viewFromInventoryID, globalVariable.ORDERINVENTORY_BYID)
            invList = New List(Of ListInventory_Data)
            For i = 0 To dtResult.Rows.Count - 1
                invList.Add(ListInventory_Data.NewListInventory(dtResult.Rows(i)("ShopID"), dtResult.Rows(i)("ShopCode"), dtResult.Rows(i)("ShopName")))
            Next i
        Catch ex As Exception
            resultText = ex.ToString
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "InventoryModule", "ListInventoryViewForSelectInventory", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function GetProperty(ByVal globalVariable As GlobalVariable) As Boolean
        Dim dtProperty As New DataTable
        dtProperty = InventorySQL.GetProperty(globalVariable.DocDBUtil, globalVariable.DocConn)
        If dtProperty.Rows.Count > 0 Then
            globalVariable.DigitForRoundingDecimal = dtProperty.Rows(0)("DigitForRoundingDecimal")
            globalVariable.CurrencySymbol = dtProperty.Rows(0)("CurrencySymbol")
            globalVariable.CurrencyCode = dtProperty.Rows(0)("CurrencyCode")
            globalVariable.CurrencyName = dtProperty.Rows(0)("CurrencyName")
            globalVariable.CurrencyFormat = dtProperty.Rows(0)("CurrencyFormat")
            globalVariable.DateFormat = dtProperty.Rows(0)("DateFormat")
            globalVariable.TimeFormat = dtProperty.Rows(0)("TimeFormat")
            globalVariable.QtyFormat = dtProperty.Rows(0)("QtyFormat")
            globalVariable.ShortDate = dtProperty.Rows(0)("ShortDate")
            globalVariable.ShortDateTime = dtProperty.Rows(0)("ShortDateTime")
            globalVariable.MaterialQtyFormat = dtProperty.Rows(0)("MaterialQtyFormat")
            globalVariable.NumericFormat = dtProperty.Rows(0)("NumericFormat")
            globalVariable.FullDateFormat = dtProperty.Rows(0)("FullDateFormat")
            globalVariable.FullDateTimeFormat = dtProperty.Rows(0)("FullDateTimeFormat")
            globalVariable.AccountingFormat = dtProperty.Rows(0)("AccountingFormat")
        End If
        Return True
    End Function
End Module
