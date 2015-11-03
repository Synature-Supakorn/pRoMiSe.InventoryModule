
Public Class Document_Data
    Public SoftwareVersion As String
    Public DocumentType As Integer
    Public DocumentTypeName As String
    Public DocumentID As Integer
    Public DocumentShopID As Integer
    Public DocumentStatus As Integer
    Public DocumentStatusName As String
    Public DocumentRefID As Integer
    Public DocumentRefShopID As Integer
    Public DocumentRefStatus As Integer
    Public DocumentDate As String
    Public DocumentNumber As String
    Public DocumentRefNumber As String
    Public DocumentInventoryID As Integer
    Public DocumentToInventoryID As Integer
    Public DocumentFromInventoryID As Integer
    Public MaxDocDetailID As Integer
    Public MaterialGroupType As String
    Public ListMaterialBy As Integer
    Public LockEditDocDetail As Boolean
    Public VendorID As Integer
    Public VendorGroupID As Integer
    Public VendorGroupShopID As Integer
    Public DueDate As String
    Public DeliveryTime As String
    Public DeliveryCost As Decimal
    Public DeliveryCostVAT As Decimal
    Public DeliveryCostNetPrice As Decimal
    Public TermOfPayment As Integer
    Public CreditDay As Integer
    Public InvoicePODate As String
    Public InvoicePODateForDisplay As String
    Public InvoicePOTotalPriceBeforeVAT As Decimal
    Public InvoicePOTotalPriceIncludeVAT As Decimal
    Public CompareAmountFromType As Integer
    Public ApproveDocStaffID As Integer
    Public ApproveDocStaffName As String
    Public TransferReceiveByStaffID As Integer
    Public InsertStaffID As Integer
    Public InsertStaffName As String
    Public UpdateStaffID As Integer
    Public UpdateStaffName As String
    Public CancelStaffID As Integer
    Public CancelStaffName As String
    Public InsertDate As String
    Public UpdateDate As String
    Public ApproveDate As String
    Public CancelDate As String
    Public thisDocumentInTemp As Boolean
    Public DocumentNote As String
    Public InvoiceRef As String
    Public ComeFromDocumentType As Integer
    Public GoToDocumentType As Integer
    Public FromTemplateID As Integer
    Public FromTemplateShopID As Integer
    Public IsAddReduceDoc As Integer
    Public StockAtDateTime As String
    Public LastTransferStock As String
    Public DefaultTaxType As Integer
    Public DocDetailList As List(Of DocumentDetail_Data)
    Public DocSummary As DocumentPriceSummary_Data
    Public MaterialNotEnoughStock As List(Of MaterialNotEnoughStock_Data)
End Class

Public Class DocumentPTT_Data
    Public SoftwareVersion As String
    Public DocumentType As Integer
    Public DocumentTypeName As String
    Public DocumentID As Integer
    Public DocumentShopID As Integer
    Public DocumentStatus As Integer
    Public DocumentStatusName As String
    Public DocumentRefID As Integer
    Public DocumentRefShopID As Integer
    Public DocumentRefStatus As Integer
    Public DocumentDate As String
    Public DocumentNumber As String
    Public DocumentRefNumber As String
    Public DocumentInventoryID As Integer
    Public DocumentToInventoryID As Integer
    Public DocumentFromInventoryID As Integer
    Public MaxDocDetailID As Integer
    Public MaterialGroupType As String
    Public ListMaterialBy As Integer
    Public LockEditDocDetail As Boolean
    Public VendorID As Integer
    Public VendorGroupID As Integer
    Public VendorGroupShopID As Integer
    Public DueDate As String
    Public TermOfPayment As Integer
    Public CreditDay As Integer
    Public InvoicePODate As String
    Public InvoicePODateForDisplay As String
    Public InvoicePOTotalPriceBeforeVAT As Decimal
    Public InvoicePOTotalPriceIncludeVAT As Decimal
    Public CompareAmountFromType As Integer
    Public ApproveDocStaffID As Integer
    Public ApproveDocStaffName As String
    Public TransferReceiveByStaffID As Integer
    Public InsertStaffID As Integer
    Public InsertStaffName As String
    Public UpdateStaffID As Integer
    Public UpdateStaffName As String
    Public CancelStaffID As Integer
    Public CancelStaffName As String
    Public InsertDate As String
    Public UpdateDate As String
    Public ApproveDate As String
    Public CancelDate As String
    Public thisDocumentInTemp As Boolean
    Public DocumentNote As String
    Public InvoiceRef As String
    Public ComeFromDocumentType As Integer
    Public GoToDocumentType As Integer
    Public FromTemplateID As Integer
    Public FromTemplateShopID As Integer
    Public CustomerDocNo As String
    Public CustomerCode As String
    Public CustomerName As String
    Public CustomerAddress As String
    Public CustomerShipTo As String
    Public CustomerBillTo As String
    Public TaxInvoiceNo As String
    Public TaxInvoiceDate As String
    Public InvoiceNo As String
    Public InvoiceDate As String
    Public SaleOrderNo As String
    Public SaleOrderDate As String
    Public PurchaseOrderNo As String
    Public PurchaseOrderDate As String
    Public DeliveryNo As String
    Public DeliveryDate As String
    Public ShippingCostTaxType As Integer
    Public DeliveryCostTaxType As Integer
    Public DeliveryCost As Decimal
    Public DeliveryCostVAT As Decimal
    Public DeliveryCostNetPrice As Decimal
    Public BusinessID As Integer
    Public PlantID As Integer
    Public GS_No As String
    Public PaymentDate As String
    Public Incoterm As String
    Public ContractNo As String
    Public ShipmentDate As String
    Public ShippingCondition As String
    Public CarrierBy As String
    Public CarrierNo As String
    Public CarrierName As String
    Public DriverName As String
    Public SealNo As String
    Public TripNo As String
    Public Contractor As String
    Public ShippingCost As Decimal
    Public ShipmentNo As String
    Public ShiftID As Integer
    Public ShiftDay As Integer
    Public ShiftNo As Integer
    Public CreateBy As String
    Public IsAddReduceDoc As Integer
    Public DefaultTaxType As Integer
    Public DocDetailList As List(Of DocumentDetail_Data)
    Public DocSummary As DocumentPriceSummary_Data
    Public DocDetailTankList As List(Of DocumentDetailTank_Data)

End Class

Public Class DocumentPriceSummary_Data
    Public SubTotal As Decimal
    Public Discount As Decimal
    Public Delivery As Decimal
    Public DeliveryNetPrice As Decimal
    Public DeliveryVAT As Decimal
    Public NetPrice As Decimal
    Public TotalVAT As Decimal
    Public GrandTotal As Decimal
End Class

Public Class DocumentDetailTank_Data
    Public DocumentID As Integer
    Public DocumentShopID As Integer
    Public DocDetailID As Integer
    Public TankID As Integer
    Public TankNo As String
    Public TankName As String
    Public ProductAmount As Decimal

    Public Shared Function NewTankData(ByVal docDetailId As Integer, ByVal documentId As String, ByVal documentShopId As String, ByVal tankId As Integer, ByVal tankNo As String, ByVal tankName As String, ByVal addAmount As Decimal) As DocumentDetailTank_Data
        Dim aData As New DocumentDetailTank_Data
        aData.DocumentID = documentId
        aData.DocumentShopID = documentShopId
        aData.DocDetailID = docDetailId
        aData.TankID = tankId
        aData.TankNo = tankNo
        aData.TankName = tankName
        aData.ProductAmount = addAmount
        Return aData
    End Function
End Class

Public Class DocumentDetail_Data

    Public DocumentID As Integer
    Public DocumentShopID As Integer
    Public DocDetailID As Integer
    Public DocDetailOrdering As Integer
    Public MaterialID As Integer
    Public MaterialCode As String
    Public MaterialName As String
    Public MaterialCode1 As String
    Public MaterialName1 As String
    Public Amount As Decimal
    Public PricePerUnit As Decimal
    Public MaterialDiscount As Decimal
    Public MaterialDiscountValue As Decimal
    Public MaterialVAT As Decimal
    Public MaterialTotalPriceBeforeDiscount As Decimal
    Public MaterialNetPrice As Decimal
    Public PricePerUnitBeforeMark As Decimal
    Public MarkUpPercent As Decimal
    Public UnitName As String
    Public UnitID As Integer
    Public UnitSmallID As Integer
    Public UnitSmallAmount As Decimal
    Public MaterialVATType As Integer
    Public RequestSmallAmount As Decimal
    Public PrepareSmallAmount As Decimal
    Public TransferSmallAmount As Decimal
    Public ROSmallAmount As Decimal
    Public DefaultInCompare As Boolean
    Public LastOrderSmallAmount As Decimal
    Public LastOrderDisplayAmount As Decimal
    Public LastOrderNetPrice As Decimal
    Public LastOrderTax As Decimal
    Public LastOrderDate As String
    Public ReferenceNetPrice As Decimal
    Public ReferenceProductTax As Decimal
    Public DiscountPercent As Decimal
    Public DiscountAmount As Decimal
    Public MaterialDiscountType As Integer
    Public CurrentStock As Decimal
    Public DiffAmount As Decimal
    Public ExtraText1 As String
    Public ExtraText2 As String
    Public MatTemp As Decimal
    Public TestTemp As Decimal
    Public TestAPI As Decimal
    Public UnitSmallName As String

    Public Shared Function AddOrUpdateDocDetailData(ByVal docDetailList As List(Of DocumentDetail_Data), ByVal docIndex As Integer,
                                                    ByVal docID As Integer, ByVal docShopID As Integer, ByVal docDetailID As Integer,
                                                    ByVal docDetailOrdering As Integer, ByVal materialID As Integer, ByVal materialCode As String,
                                                    ByVal materialName As String, ByVal materialVATType As Integer, ByVal unitName As String,
                                                    ByVal unitID As Integer, ByVal unitSmallID As Integer, ByVal amount As Decimal, ByVal pricePerUnit As Decimal,
                                                    ByVal totalPriceBeforeDiscount As Decimal, ByVal materialDiscount As Decimal, ByVal discountPercent As Decimal,
                                                    ByVal discountAmount As Decimal, ByVal materialVAT As Decimal, ByVal materialNetPrice As Decimal,
                                                    ByVal unitSmallAmount As Decimal, ByVal rqSmallAmount As Decimal, ByVal prepareSmallAmount As Decimal,
                                                    ByVal transferSmallAmount As Decimal, ByVal roSmallAmount As Decimal, ByVal defuaultInCompare As Integer,
                                                    ByVal lastOrderSmallAmount As Decimal, ByVal lastOrderDisplayAmount As Decimal, ByVal lastOrderNetPrice As Decimal,
                                                    ByVal lastOrderTax As Decimal, ByVal lastOrderDate As Date, ByVal refNetPrice As Decimal, ByVal refProductTax As Decimal,
                                                    ByVal stockAmount As Decimal, ByVal diffStockAmount As Decimal, ByVal remark As String, ByVal api60F As String,
                                                    ByVal matTemp As Decimal, ByVal testTemp As Decimal, ByVal testApi As Decimal, ByVal unitSmallName As String,
                                                    ByVal materialCode1 As String, ByVal materialName1 As String)

        Dim docData As DocumentDetail_Data
        If docIndex = -1 Then
            docData = New DocumentDetail_Data
            docDetailList.Add(docData)
        Else
            docData = docDetailList(docIndex)
        End If
        docData.DocumentID = docID
        docData.DocumentShopID = docShopID
        docData.DocDetailID = docDetailID
        docData.DocDetailOrdering = docDetailOrdering
        docData.MaterialID = materialID
        docData.MaterialCode = materialCode
        docData.MaterialName = materialName
        docData.MaterialVATType = materialVATType
        docData.UnitName = unitName
        docData.UnitID = unitID
        docData.UnitSmallID = unitSmallID
        docData.Amount = amount
        docData.PricePerUnit = pricePerUnit
        docData.MaterialDiscount = materialDiscount
        docData.DiscountPercent = discountPercent
        docData.DiscountAmount = discountAmount
        If discountPercent > 0 Then
            docData.MaterialDiscountValue = discountPercent
        End If
        If discountAmount > 0 Then
            docData.MaterialDiscountValue = discountAmount
        End If
        docData.MaterialVAT = materialVAT
        docData.MaterialNetPrice = materialNetPrice
        docData.MaterialTotalPriceBeforeDiscount = totalPriceBeforeDiscount
        docData.UnitSmallAmount = unitSmallAmount
        docData.RequestSmallAmount = rqSmallAmount
        docData.PrepareSmallAmount = prepareSmallAmount
        docData.TransferSmallAmount = transferSmallAmount
        docData.ROSmallAmount = roSmallAmount
        If defuaultInCompare = 1 Then
            docData.DefaultInCompare = True
        Else
            docData.DefaultInCompare = False
        End If
        docData.LastOrderSmallAmount = lastOrderSmallAmount
        docData.LastOrderDisplayAmount = lastOrderDisplayAmount
        docData.LastOrderNetPrice = lastOrderNetPrice
        docData.LastOrderTax = lastOrderTax
        docData.LastOrderDate = lastOrderDate
        docData.ReferenceNetPrice = refNetPrice
        docData.ReferenceProductTax = refProductTax
        If docData.DiscountAmount > 0 And docData.DiscountPercent = 0 Then
            docData.MaterialDiscountType = 1
        Else
            docData.MaterialDiscountType = 2
        End If
        docData.CurrentStock = stockAmount
        docData.DiffAmount = (amount - stockAmount)
        docData.ExtraText1 = remark
        docData.ExtraText2 = api60F
        docData.MatTemp = matTemp
        docData.TestTemp = testTemp
        docData.TestAPI = testApi
        docData.UnitSmallName = unitSmallName
        docData.MaterialCode1 = materialCode1
        docData.MaterialName1 = materialName1
        Return docData
    End Function
   
End Class

Public Class SearchDocumentResult_Data
    Public DocumentID As Integer
    Public DocumentShopID As Integer
    Public DocumentStatus As Integer
    Public DocumentStatusName As String
    Public DocumentType As Integer
    Public DocumentTypeName As String
    Public DocumentRefID As Integer
    Public DocumentRefShopID As Integer
    Public DocumentRefStatus As Integer
    Public DocumentRefDocumentType As Integer
    Public DocumentRefDocumentTypeName As String
    Public DocumentDate As String
    Public DueDate As String
    Public DocumentNumber As String
    Public DocumentRefNumber As String
    Public ToDocumentNumber As String
    Public DocumentInventoryID As Integer
    Public DocumentInvetoryName As String
    Public DocumentToInventoryID As Integer
    Public DocumentToInventoryName As String
    Public DocumentFromInventoryID As Integer
    Public DocumentFromInventoryName As String
    Public DocumentNote As String
    Public InvoiceRef As String
    Public DocumentReceiveBy As Integer
    Public VendorID As Integer
    Public VendorGroupID As Integer
    Public VendorShopID As Integer
    Public VendorCode As String
    Public VendorName As String
    Public InsertStaffName As String
    Public UpdateStaffName As String
    Public CancelStaffName As String
    Public ApproveStaffName As String
    Public SubTotal As Decimal
    Public TotalDiscount As Decimal
    Public TotalVat As Decimal
    Public NetPrice As Decimal
    Public GrandTotal As Decimal
    Public BusinessPlace As String
    Public TaxInvoiceNo As String
    Public TaxInvoiceDate As String

    Public Shared Function NewSearchDocumentResult(ByVal docID As Integer, ByVal docShopID As Integer, ByVal docStatus As Integer, ByVal docStatusName As String,
    ByVal docReceiveBy As Integer, ByVal docType As Integer, ByVal docTypeName As String,
    ByVal docRefID As Integer, ByVal docRefShopID As Integer, ByVal docRefStatus As Integer, ByVal docRefDocType As Integer,
    ByVal docRefDocTypeName As String, ByVal docDate As Date, ByVal dueDate As DateTime, ByVal docNumber As String, ByVal docRefNumber As String,
    ByVal docInvID As Integer, ByVal docInvName As String, ByVal docToInvID As String, ByVal docToInvName As String, ByVal docFromInvID As Integer,
    ByVal docFromInvName As String, ByVal docNote As String, ByVal invoiceRef As String,
    ByVal vendorID As Integer, ByVal vendorGroupID As Integer, ByVal vendorShopID As Integer, ByVal vendorCode As String,
    ByVal vendorName As String, ByVal subTotal As Decimal, ByVal totalDiscount As Decimal, ByVal totalVat As Decimal,
    ByVal netprice As Decimal, ByVal grandTotal As Decimal, ByVal insertStaffName As String, ByVal updateStaffName As String,
    ByVal approveStaffName As String, ByVal cancelStaffName As String, ByVal businessPlace As String, ByVal TaxInvoiceNo As String,
    ByVal TaxInvoiceDate As Date) As SearchDocumentResult_Data

        Dim nData As New SearchDocumentResult_Data
        nData.DocumentID = docID
        nData.DocumentShopID = docShopID
        nData.DocumentStatus = docStatus
        nData.DocumentStatusName = docStatusName
        nData.DocumentReceiveBy = docReceiveBy
        nData.DocumentType = docType
        nData.DocumentTypeName = docTypeName

        nData.DocumentRefID = docRefID
        nData.DocumentRefShopID = docRefShopID
        nData.DocumentRefStatus = docRefStatus
        nData.DocumentRefDocumentType = docRefDocType
        nData.DocumentRefDocumentTypeName = docRefDocTypeName

        nData.DocumentDate = docDate.ToString("yyyy-MM-dd")
        nData.DueDate = dueDate.ToString("yyyy-MM-dd HH:mm:ss")
        nData.DocumentNumber = docNumber
        nData.DocumentRefNumber = docRefNumber
        nData.DocumentInventoryID = docInvID
        nData.DocumentInvetoryName = docInvName
        nData.DocumentToInventoryID = docToInvID
        nData.DocumentToInventoryName = docToInvName
        nData.DocumentFromInventoryID = docInvID
        nData.DocumentFromInventoryName = docFromInvName
        nData.DocumentNote = docNote
        nData.InvoiceRef = invoiceRef

        nData.VendorID = vendorID
        nData.VendorGroupID = vendorGroupID
        nData.VendorShopID = vendorShopID
        nData.VendorCode = vendorCode
        nData.VendorName = vendorName

        nData.InsertStaffName = insertStaffName
        nData.UpdateStaffName = updateStaffName
        nData.CancelStaffName = cancelStaffName
        nData.ApproveStaffName = approveStaffName

        nData.SubTotal = Format(subTotal, "#,##0.00")
        nData.TotalDiscount = Format(totalDiscount, "#,##0.00")
        nData.TotalVat = Format(totalVat, "#,##0.00")
        nData.NetPrice = Format(netprice, "#,##0.00")
        nData.GrandTotal = Format(grandTotal, "#,##0.00")

        nData.BusinessPlace = businessPlace
        nData.TaxInvoiceNo = TaxInvoiceNo
        nData.TaxInvoiceDate = TaxInvoiceDate.ToString("yyyy-MM-dd")
        Return nData
    End Function

End Class

Public Class AddReduceDocumentType_Data
    Public DocumentTypeID As Integer
    Public DocumentTypeHeader As String
    Public DocumentTypeName As String
    Public MovementInStock As Integer
    Public Const MOVEMENT_ADD As Integer = 1
    Public Const MOVEMENT_REDUCE As Integer = -1

    Public Shared Function NewAddReduceDocData(ByVal docTypeID As Integer, ByVal docTypeHeader As String, ByVal docTypeName As String, ByVal movementInStock As Integer) As AddReduceDocumentType_Data
        Dim aData As New AddReduceDocumentType_Data
        aData.DocumentTypeID = docTypeID
        aData.DocumentTypeHeader = docTypeHeader
        aData.DocumentTypeName = docTypeName
        aData.MovementInStock = movementInStock
        Return aData
    End Function

    Public Shared Function ListDocTypeAddReduce() As List(Of AddReduceDocumentType_Data)
        Dim taxtypeList As New List(Of AddReduceDocumentType_Data)
        taxtypeList.Add(AddReduceDocumentType_Data.NewAddReduceDocData(1, "", "ปรับเพิ่มสต๊อก", MOVEMENT_ADD))
        taxtypeList.Add(AddReduceDocumentType_Data.NewAddReduceDocData(2, "", "ปรับลดสต๊อก", MOVEMENT_REDUCE))
        Return taxtypeList
    End Function

End Class

