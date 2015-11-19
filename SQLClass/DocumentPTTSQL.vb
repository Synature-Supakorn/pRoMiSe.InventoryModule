Imports pRoMiSe.DBHelper
Imports pRoMiSe.Utilitys.Utilitys
Imports System.Data.SqlClient
Imports System.Text

Module DocumentPTTSQL


    Friend Function DeleteDocumentDetailTank(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                            ByVal documentShopID As Integer, ByVal docDetailID As Integer) As String
        Dim strSQL As String = ""
        strSQL = "Delete From DocDetailTank " &
                 " Where DocDetailId=" & docDetailID & " AND DocumentID=" & documentID & " AND ShopID=" & documentShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function DeleteDocumentDetailTank(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                            ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal tankID As Integer) As String
        Dim strSQL As String = ""
        strSQL = "Delete From DocDetailTank " &
                 " Where DocDetailId=" & docDetailID & " AND DocumentID=" & documentID & " AND ShopID=" & documentShopID & " AND TankID=" & tankID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function GetDocumentDetailTank(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentID As Integer, ByVal documentShopID As Integer) As DataTable
        Dim strSQL As String = ""
        strSQL = "select * from DocDetailTank " &
                 " Where DocumentID=" & documentID & " AND ShopID=" & documentShopID
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetDocumentDetailTank(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal docDetailId As Integer, ByVal documentID As Integer, ByVal documentShopID As Integer) As DataTable
        Dim strSQL As String = ""
        strSQL = "select * from DocDetailTank " &
                 " Where DocumentID=" & documentID & " AND ShopID=" & documentShopID & " AND DocDetailId=" & docDetailId
        Return dbUtil.List(strSQL, connection)
    End Function


    Friend Function GetTankDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection) As DataTable
        Dim strSQL As String = ""
        strSQL = "select * from ENABLERDB.[dbo].tanks"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function InsertDocumentDetailTank(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                             ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal tankID As Integer, ByVal tankNo As String,
                                             ByVal tankName As String, ByVal materialAmount As Decimal) As String
        Dim strSQL As String = ""
        strSQL = "Insert Into DocDetailTank(DocDetailId,DocumentID,ShopID,TankID,TankNo,TankName,ProductAmount)" &
                 "values(" & docDetailID & "," & documentID & "," & documentShopID & "," & tankID & ",'" & tankNo & "','" & Trim(tankName) & "'," & materialAmount & ") "
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function
  
    Friend Function GetMaterialOil(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection) As DataTable
        Dim strSQL As String = ""
        strSQL = "select * from Materials where MaterialIDRef<>0 and isshowinpos=1"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function InsertDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                               ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal materialID As Integer, ByVal materialAmount As Decimal,
                                               ByVal percentDiscount As Decimal, ByVal amountDiscount As Decimal, ByVal pricePerUnit As Decimal, ByVal materialTax As Decimal,
                                               ByVal taxType As Integer, ByVal materialUnitSmallID As Integer, ByVal selectUnitID As Integer, ByVal selectUnitName As String,
                                               ByVal unitSmallAmount As Decimal, ByVal materialNetPrice As Decimal, ByVal materialCode As String,
                                               ByVal materialName As String, ByVal supplierMaterialCode As String, ByVal supplierMaterialName As String,
                                               ByVal remark As String, ByVal Api60F As String, ByVal unitSmallName As String, ByVal unitSAP As Integer) As String
        Dim strSQL As String = ""
        strSQL = "Insert INTO DocDetail(DocDetailID, DocumentID, ShopID, ProductID, ProductUnit, ProductAmount, " &
                 "ProductDiscount, ProductDiscountAmount, ProductPricePerUnit,ProductTax, ProductTaxType, UnitID, " &
                 "UnitName, UnitSmallAmount, ProductNetPrice,ProductCode,ProductName,SupplierMaterialCode,SupplierMaterialName,ExtraText1,ExtraText2,ExtraText3,AdjustLinkGroup) " &
                 "Values(" & docDetailID & ", " & documentID & ", " & documentShopID & ", " & materialID & ", " & materialUnitSmallID & ", " & materialAmount &
                 "," & percentDiscount & ", " & amountDiscount & "," & pricePerUnit & ", " & materialTax & ", " & taxType & ", " & selectUnitID &
                 ",'" & selectUnitName & "', " & unitSmallAmount & "," & materialNetPrice & ",'" & materialCode & "','" & materialName &
                 "','" & supplierMaterialCode & "','" & supplierMaterialName & "','" & ReplaceSuitableStringForSQL(remark) & "','" & ReplaceSuitableStringForSQL(Api60F) & "','" & ReplaceSuitableStringForSQL(unitSmallName) & "'," & unitSAP & ")"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocumentFormPTT(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction,
     ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As String,
     ByVal customerDocNo As String, customerCode As String, ByVal customerName As String, ByVal customerAddress As String, ByVal customerShipTo As String, ByVal customerBillTo As String,
     ByVal taxInvoiceNo As String, ByVal taxInvoiceDate As String, ByVal invoiceNo As String, ByVal invoiceDate As String,
     ByVal saleOrderNo As String, ByVal saleOrderDate As String, ByVal purchaseOrderNo As String, ByVal purchaseOrderDate As String,
     ByVal deliveryNo As String, ByVal deliveryDate As String, ByVal businessId As Integer, ByVal plantId As Integer,
     ByVal paymentDate As String, ByVal incoterm As String, ByVal contractNo As String, ByVal shipmentDate As String, ByVal shippingCondition As String,
     ByVal carrierBy As String, ByVal carrierName As String, ByVal driverName As String, ByVal sealNo As String, ByVal tripNo As String,
     ByVal contractor As String, ByVal transferTaxClass As Integer, ByVal transferTotal As Decimal, ByVal transferVAT As Decimal, ByVal transferNetPrice As Decimal,
     ByVal ShipmentNo As String, ByVal ShiftID As Integer, ByVal ShiftDay As Integer, ByVal ShiftNo As Integer, ByVal GSNO As String,
     ByVal updateDate As String, ByVal updateBy As Integer, ByVal createBy As String) As Integer

        Dim strSQL As String
        strSQL = "UPDATE  Document SET " &
                "DocumentDate  =" & documentDate &
                ", CustomerDocNo='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(customerDocNo) & "'" &
                ", CustomerCode  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(customerCode) & "'" &
                ", CustomerName ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(customerName) & "'" &
                ", CustomerAddress  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(customerAddress) & "'" &
                ", ShipToOrDestinationPort  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(customerShipTo) & "'" &
                ", BillTo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(customerBillTo) & "'" &
                ", TaxInvoiceNo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(taxInvoiceNo) & "'" &
                ", TaxInvoiceDate  =" & taxInvoiceDate &
                ", InvoiceNo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(invoiceNo) & "'" &
                ", InvoiceDate  =" & invoiceDate &
                ", SaleOrderNo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(saleOrderNo) & "'" &
                ", SaleOrderDate  =" & saleOrderDate &
                ", InvoicePONO ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(purchaseOrderNo) & "'" &
                ", InvoicePODate  =" & purchaseOrderDate &
                ", DeliveryNo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(deliveryNo) & "'" &
                ", DeliveryDate  =" & deliveryDate &
                ", BusinessPlace  =" & businessId &
                ", Plant  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(plantId) & "'" &
                ", PaymentDate  = " & paymentDate &
                ", Incoterm  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(incoterm) & "'" &
                ", ContractNo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(contractNo) & "'" &
                ", ShipmentDate  =" & shipmentDate &
                ", ShippingCondition  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(shippingCondition) & "'" &
                ", CarrierBy  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(carrierBy) & "'" &
                ", CarrierNo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(carrierName) & "'" &
                ", DriverName  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(driverName) & "'" &
                ", SealNo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(sealNo) & "'" &
                ", TripNo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(tripNo) & "'" &
                ", TransferTaxClass  =" & transferTaxClass &
                ", TransferTotal  =" & transferTotal &
                ", TransferVAT  =" & transferVAT &
                ", TransferNetPrice  =" & transferNetPrice &
                ", ShipmentNo  ='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(ShipmentNo) & "'" &
                ", ShiftID  =" & ShiftID &
                 ", ShiftDay  =" & ShiftDay &
                ", ShiftNo  =" & ShiftNo &
                ", GSNO='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(GSNO) & "'" &
                ",UpdateBy=" & updateBy &
                ",CreateBy='" & createBy & "'" &
                ",Remark='" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(contractor) & "'" &
                 " Where DocumentID = " & documentID & " AND ShopID = " & documentShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                                 ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal materialID As Integer, ByVal materialAmount As Decimal,
                                                 ByVal percentDiscount As Decimal, ByVal amountDiscount As Decimal, ByVal pricePerUnit As Decimal, ByVal materialTax As Decimal,
                                                 ByVal taxType As Integer, ByVal materialUnitSmallID As Integer, ByVal selectUnitID As Integer, ByVal selectUnitName As String,
                                                 ByVal unitSmallAmount As Decimal, ByVal materialNetPrice As Decimal, ByVal materialCode As String,
                                                 ByVal materialName As String, ByVal supplierMaterialCode As String, ByVal supplierMaterialName As String,
                                                 ByVal remark As String, ByVal Api60F As String, ByVal matTemp As Decimal, ByVal testTemp As Decimal, ByVal testApi As Decimal,
                                                 ByVal unitSmallName As String, ByVal unitSAP As Integer) As Integer
        Dim strSQL As String
        strSQL = "Update DocDetail " &
                 "Set ProductID = " & materialID & ", ProductUnit = " & materialUnitSmallID &
                 ", ProductAmount = " & materialAmount & ", ProductDiscount = " & percentDiscount &
                 ", ProductDiscountAmount = " & amountDiscount & ", ProductPricePerUnit = " & pricePerUnit &
                 ", ProductTax = " & materialTax & ", ProductTaxType = " & taxType &
                 ", UnitID = " & selectUnitID & ", UnitName = '" & selectUnitName &
                 "',UnitSmallAmount = " & unitSmallAmount & ", ProductNetPrice = " & materialNetPrice &
                 ", ProductCode='" & materialCode & "',ProductName='" & materialName & "',SupplierMaterialCode='" & supplierMaterialCode & _
                 "',SupplierMaterialName='" & supplierMaterialName & "',ExtraText1='" & Trim(ReplaceSuitableStringForSQL(remark)) & "',ExtraText2='" & Trim(ReplaceSuitableStringForSQL(Api60F)) & "'" & _
                 ",ExtraValue1=" & matTemp & ",ExtraValue2=" & testTemp & ",ExtraValue3=" & testApi & ",ExtraText3='" & Trim(ReplaceSuitableStringForSQL(unitSmallName)) & "',AdjustLinkGroup=" & unitSAP & _
                 " Where DocDetailID = " & docDetailID & " AND DocumentID = " & documentID & " AND ShopID = " & documentShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocumentDetailTank(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                            ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal tankID As Integer, ByVal materialAmount As Decimal) As String
        Dim strSQL As String = ""
        strSQL = "Update DocDetailTank Set ProductAmount=" & materialAmount &
                 " Where DocDetailId=" & docDetailID & " AND DocumentID=" & documentID & " AND ShopID=" & documentShopID & " AND TankID=" & tankID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function SumReceiveOilToDocDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer, ByVal documentShopId As Integer) As Integer
        Dim strSQL As String = ""
        strSQL = "update DocDetail Set DocDetail.ROSmallAmount = tank.Qty" &
                 " from (select DocDetailId,DocumentID,ShopID,sum(ProductAmount) As Qty from DocDetailTank " &
                 " where documentid = " & documentID & " And shopid = " & documentShopId &
                 " group by DocDetailId,DocumentID,ShopID) Tank " &
                 " Where DocDetail.DocDetailId = Tank.DocDetailId And DocDetail.DocumentID = Tank.DocumentID And DocDetail.ShopID = Tank.ShopID"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

End Module
