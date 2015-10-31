Imports pRoMiSe.Utilitys.Utilitys
Imports System.Data.SqlClient

Module DirectReceiveOrderModule

    Friend Function AddDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopID As Integer, ByVal materialID As Integer,
                                ByVal addAmount As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal, ByVal discountAmount As Decimal,
                                ByVal discountPercent As Decimal, ByVal materialVATType As Integer, ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        Dim dtMaterialUnit As DataTable
        Dim selDocDetailID As Integer
        Dim selTax, selUnitSmallAmount As Decimal
        Dim selUnitSmallID, selUnitID As Integer
        Dim selUnitName As String
        Dim selDiscountPrice As Decimal
        Dim selMaterialNetPrice, selTotalPriceBeforeDiscount As Decimal
        Dim selMaterialCode As String = ""
        Dim selMaterialName As String = ""
        Dim selMaterialSupplierCode As String = ""
        Dim selMaterialSupplierName As String = ""
        Dim dtProperty As New DataTable
        Dim digitDecimal As Integer = 2
        dtProperty = InventorySQL.GetProperty(globalVariable.DocDBUtil, globalVariable.DocConn)
        If dtProperty.Rows.Count > 0 Then
            digitDecimal = dtProperty.Rows(0)("DigitForRoundingDecimal")
        End If

        dtMaterialUnit = MaterialSQL.GetMaterialDetailAndUnitRatio(globalVariable.DocDBUtil, globalVariable.DocConn, materialID, materialUnitLargeID, False)
        If dtMaterialUnit.Rows.Count = 0 Then
            resultText = "ไม่พบวัตถุดิบที่เลือก"
            Return False
        End If
        selUnitID = dtMaterialUnit.Rows(0)("SelectUnitID")
        selUnitName = dtMaterialUnit.Rows(0)("UnitLargeName")
        selUnitSmallID = dtMaterialUnit.Rows(0)("UnitSmallID")

        If Not IsDBNull(dtMaterialUnit.Rows(0)("MaterialCode")) Then
            selMaterialCode = dtMaterialUnit.Rows(0)("MaterialCode")
        End If
        If Not IsDBNull(dtMaterialUnit.Rows(0)("MaterialName")) Then
            selMaterialName = dtMaterialUnit.Rows(0)("MaterialName")
        End If
        If Not IsDBNull(dtMaterialUnit.Rows(0)("MaterialCode1")) Then
            selMaterialSupplierCode = dtMaterialUnit.Rows(0)("MaterialCode1")
        End If
        If Not IsDBNull(dtMaterialUnit.Rows(0)("MaterialName1")) Then
            selMaterialSupplierName = dtMaterialUnit.Rows(0)("MaterialName1")
        End If

        selUnitSmallAmount = Format((addAmount * dtMaterialUnit.Rows(0)("UnitSmallRatio")) / dtMaterialUnit.Rows(0)("UnitLargeRatio"), "0.0000")
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            selDocDetailID = DocumentSQL.GetMaxDocDetailID(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)
            DocumentModule.CalculateDocDetailAllPrice(globalVariable, addAmount, pricePerUnit, discountPercent, discountAmount, materialVATType,
                                                      selTotalPriceBeforeDiscount, selDiscountPrice, selTax, selMaterialNetPrice)
            DocumentSQL.InsertDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, selDocDetailID, materialID, addAmount,
                                              FormatDecimal(discountPercent, digitDecimal), FormatDecimal(discountAmount, digitDecimal), FormatDecimal(pricePerUnit, digitDecimal), FormatDecimal(selTax, digitDecimal), materialVATType, selUnitSmallID, selUnitID, selUnitName,
                                              selUnitSmallAmount, FormatDecimal(selMaterialNetPrice, digitDecimal), selMaterialCode, selMaterialName, selMaterialSupplierCode, selMaterialSupplierName)
            DocumentSQL.UpdateDocSummaryIntoDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DirectReceiveOrderModule", "AddDocDetail", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function ApproveDocument(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer, ByRef resultText As String) As Boolean
        Dim dbTrans As SqlTransaction
        Dim updateDate As DateTime
        Dim strUpdateDate As String
        Dim vendorId As Integer = 0
        Dim dtDoc As New DataTable

        dtDoc = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId, globalVariable.DocLangID)
        updateDate = Now
        strUpdateDate = FormatDateTime(updateDate)
        dbTrans = globalVariable.DocConn.BeginTransaction
        Try

            DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, globalVariable.StaffID, strUpdateDate)
            DocumentSQL.DeleteZeroCompareAmountMaterialInDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
            DocumentModule.UpdateMaterialDefaultPrice(globalVariable, dbTrans, documentId, documentShopId, dtDoc.Rows(0)("VendorId"))
            MaterialSQL.AutoAddDailyStockMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
            MaterialSQL.AutoAddMonthlyStockMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
            MaterialSQL.AutoAddWeeklyStockMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DirectReceiveOrderModule", "ApproveDocument", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function CancelDocument(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer, ByRef resultText As String) As Boolean

        Dim dtDocument As New DataTable
        Dim dbTrans As SqlTransaction
        Dim strUpdateDate As String

        dtDocument = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId, globalVariable.DocLangID)
        If dtDocument.Rows.Count = 0 Then
            resultText = "ไม่พบเอกสารที่ต้องการยกเลิก"
            Return False
        End If

        If CheckValidDocumentForCancelDocument(globalVariable, dtDocument.Rows(0)("documentStatus"), dtDocument.Rows(0)("ShopId"), dtDocument.Rows(0)("DocumentDate"), resultText) = False Then
            Return False
        End If

        strUpdateDate = FormatDateTime(Now)
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try

            DocumentSQL.CancelDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, globalVariable.StaffID, strUpdateDate)
            If DocumentSQL.DocumentIsAlreadyReferTo(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, dtDocument.Rows(0)("DocumentIDRef"),
                                                    dtDocument.Rows(0)("DocumentIDRefShopID"), dtDocument.Rows(0)("DocumentTypeID"),
                                                    documentId, documentShopId) = False Then

                DocumentSQL.UpdateDocumentStatus(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, dtDocument.Rows(0)("DocumentIDRef"),
                                                 dtDocument.Rows(0)("DocumentIDRefShopID"), globalVariable.DOCUMENTSTATUS_APPROVE,
                                                 strUpdateDate, globalVariable.StaffID)
            End If

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.Message
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DirectReceiveOrderModule", "CancelDocument", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function SearchDocument(ByVal globalVariable As GlobalVariable, ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date,
                                   ByVal searchInventoryID As Integer, ByVal vendorID As Integer, ByVal vendorGroupID As Integer,
                                   ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean

        Dim dtResult, dtDocStatus As DataTable
        Dim strFromDate, strToDate As String
        If startDate = Date.MinValue Then
            strFromDate = ""
        Else
            strFromDate = FormatDate(startDate)
        End If
        If endDate = Date.MinValue Then
            strToDate = ""
        Else
            strToDate = FormatDate(endDate)
        End If
        Try
            dtResult = DocumentSQL.SearchDocument(globalVariable.DocDBUtil, globalVariable.DocConn, globalVariable.DOCUMENTTYPE_DIRECTRO, strFromDate, strToDate,
                                                       documentStatus, searchInventoryID, vendorID, vendorGroupID, globalVariable.DocLangID)
            dtDocStatus = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
            If dtResult.Rows.Count > 0 Then
                docList = DocumentModule.InsertResultDataIntoList(globalVariable, globalVariable.DOCUMENTTYPE_DIRECTRO, dtResult, dtDocStatus)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DirectReceiveOrderModule", "SearchDocument", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function SaveDocumentDataIntoDB(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal inventoryID As Integer, ByVal documenttypeId As Integer,
                                           ByVal documentDate As Date, ByVal vendorID As Integer, ByVal vendorGroupID As Integer, ByVal documentNote As String,
                                           ByVal invoiceReference As String, ByVal termOfPayment As Integer, ByVal creditDay As Integer, ByVal deliveryCost As Decimal,
                                           ByVal dueDate As DateTime, ByRef resultText As String) As Boolean

        Dim strDocDate, strDueDate, strUpdateDate As String
        Dim updateDate As DateTime
        Dim newSend As Integer = 0
        Dim dbTrans As SqlTransaction
        Dim shopVAT As Decimal
        Dim dtVendor As New DataTable
        Dim defaultTaxType As Integer = 0
        Dim deliveryTax, deliveryNetPrice As Decimal

        dtVendor = VendorSQL.GetVendorDetail(globalVariable.DocDBUtil, globalVariable.DocConn, vendorID, globalVariable.DocLangID)
        If dtVendor.Rows.Count > 0 Then
            defaultTaxType = dtVendor.Rows(0)("defaultTaxType")
        End If
        shopVAT = globalVariable.DefaultShopVAT
        strDocDate = FormatDate(documentDate)
        updateDate = Now
        strUpdateDate = FormatDateTime(updateDate)
        newSend = 0
        If dueDate <> Date.MinValue Then
            strDueDate = FormatDateTime(dueDate)
        Else
            strDueDate = "NULL"
        End If
       
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentSQL.UpdateDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, inventoryID, inventoryID, vendorID, vendorGroupID, strDocDate, Trim(documentNote), Trim(invoiceReference), termOfPayment, creditDay, strDueDate, shopVAT, strUpdateDate, globalVariable.StaffID)
            DocumentModule.CalculateDeliveryCost(globalVariable, deliveryCost, defaultTaxType, deliveryTax, deliveryNetPrice)
            DocumentSQL.UpdateDocumentDeliveryCost(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, inventoryID, defaultTaxType, FormatDecimal(deliveryCost, globalVariable.DigitForRoundingDecimal), FormatDecimal(deliveryTax, globalVariable.DigitForRoundingDecimal), FormatDecimal(deliveryNetPrice, globalVariable.DigitForRoundingDecimal))
            DocumentSQL.UpdateDocumentStatus(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, inventoryID, globalVariable.DOCUMENTSTATUS_WORKING, strUpdateDate, globalVariable.StaffID)
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DirectReceiveOrderModule", "SaveDocumentDataIntoDB", "99", ex.ToString)
            Return False
        End Try

        resultText = ""
        Return True
    End Function

    Friend Function UpdateDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer,
                                    ByVal materialID As Integer, ByVal addAmount As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal,
                                    ByVal discountAmount As Decimal, ByVal discountPercent As Decimal, ByVal materialVATType As Integer, ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        Dim dtMaterialUnit As DataTable
        Dim selTax, selUnitSmallAmount As Decimal
        Dim selUnitSmallID, selUnitID As Integer
        Dim selUnitName As String
        Dim selDiscountPrice As Decimal
        Dim selMaterialNetPrice, selTotalPriceBeforeDiscount As Decimal
        Dim selMaterialCode As String = ""
        Dim selMaterialName As String = ""
        Dim selMaterialSupplierCode As String = ""
        Dim selMaterialSupplierName As String = ""

        dtMaterialUnit = MaterialSQL.GetMaterialDetailAndUnitRatio(globalVariable.DocDBUtil, globalVariable.DocConn, materialID, materialUnitLargeID, False)
        If dtMaterialUnit.Rows.Count = 0 Then
            resultText = "ไม่พบวัตถุดิบที่เลือก"
            Return False
        End If
        selUnitID = dtMaterialUnit.Rows(0)("SelectUnitID")
        selUnitName = dtMaterialUnit.Rows(0)("UnitLargeName")
        selUnitSmallID = dtMaterialUnit.Rows(0)("UnitSmallID")
        selMaterialCode = dtMaterialUnit.Rows(0)("MaterialCode")
        selMaterialName = dtMaterialUnit.Rows(0)("MaterialName")
        selUnitSmallAmount = Format((addAmount * dtMaterialUnit.Rows(0)("UnitSmallRatio")) / dtMaterialUnit.Rows(0)("UnitLargeRatio"), "0.0000")
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentModule.CalculateDocDetailAllPrice(globalVariable, addAmount, pricePerUnit, discountPercent, discountAmount, materialVATType,
                                                      selTotalPriceBeforeDiscount, selDiscountPrice, selTax, selMaterialNetPrice)
            DocumentSQL.UpdateDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, docDetailId, materialID, addAmount,
                                             discountPercent, discountAmount, pricePerUnit, selTax, materialVATType, selUnitSmallID, selUnitID, selUnitName,
                                             selUnitSmallAmount, selMaterialNetPrice, selMaterialCode, selMaterialName, selMaterialSupplierCode, selMaterialSupplierName)
            DocumentSQL.UpdateDocSummaryIntoDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DirectReceiveOrderModule", "UpdateDocDetail", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function DeleteDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer, ByVal strDocDetailId As String,
                                    ByRef resultText As String) As Boolean

        Dim i, j As Integer
        Dim dbTrans As SqlTransaction
        Dim docDetailId() As Integer

        If strDocDetailId.IndexOf(",") >= 0 Then
            Dim arrId = strDocDetailId.Split(",")
            ReDim docDetailId(-1)
            For n As Integer = 0 To arrId.Length - 1
                ReDim Preserve docDetailId(docDetailId.Length)
                docDetailId(docDetailId.Length - 1) = arrId(n)
            Next
        Else
            ReDim docDetailId(-1)
            ReDim Preserve docDetailId(docDetailId.Length)
            docDetailId(docDetailId.Length - 1) = strDocDetailId

        End If
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            For i = 0 To docDetailId.Count - 1
                DocumentSQL.DeleteDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, docDetailId(i))
            Next i
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DirectReceiveOrderModule", "DeleteDocDetail", "99", ex.ToString)
            Return False
        End Try

        resultText = ""
        Return True
    End Function

End Module
