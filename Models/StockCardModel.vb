Public Class StockCard_Data
    Public MaterialGroupId As Integer
    Public MaterialGroupCode As String
    Public MaterialGroupName As String
    Public MaterialId As Integer
    Public MaterialCode As String
    Public MaterialName As String
    Public UnitCost As Decimal
    Public BeginningAmount As Decimal
    Public ReceiveAmount As Decimal
    Public TransferInAmount As Decimal
    Public TransferOutAmount As Decimal
    Public AdjustAmount As Decimal
    Public DestroyAmount As Decimal
    Public SaleAmount As Decimal
    Public OnhandAmount As Decimal
    Public EndingAmount As Decimal
    Public VarianceAmount As Decimal
    Public UnitName As String

    Public Shared Function NewListVendor(ByVal materialGroupId As Integer, ByVal materialGroupCode As Integer, ByVal materialGroupName As String,
                                         ByVal materailId As Integer, ByVal materialCode As String, ByVal materialName As String, ByVal unitCost As Decimal,
                                         ByVal beginningAmount As Decimal, ByVal receiveAmount As Decimal, ByVal transferInAmount As Decimal,
                                         ByVal transferOutAmount As Decimal, ByVal AdjustAmount As Decimal, ByVal saleAmount As Decimal, ByVal OnhandAmount As Decimal,
                                         ByVal endingAmount As Decimal, ByVal varianceAmount As Decimal, ByVal DestroyAmount As Decimal) As StockCard_Data

        Dim stockData As New StockCard_Data
        stockData.MaterialGroupId = materialGroupId
        stockData.MaterialGroupCode = materialGroupCode
        stockData.MaterialGroupName = materialGroupName
        stockData.MaterialId = materailId
        stockData.MaterialCode = materialCode
        stockData.MaterialName = materialName
        stockData.UnitCost = unitCost
        stockData.BeginningAmount = beginningAmount
        stockData.ReceiveAmount = receiveAmount
        stockData.TransferInAmount = transferInAmount
        stockData.TransferOutAmount = transferOutAmount
        stockData.SaleAmount = saleAmount
        stockData.AdjustAmount = AdjustAmount
        stockData.OnhandAmount = OnhandAmount
        stockData.EndingAmount = endingAmount
        stockData.VarianceAmount = varianceAmount
        stockData.DestroyAmount = DestroyAmount
        Return stockData
    End Function

End Class
