Imports pRoMiSe.Utilitys.Utilitys
Imports System.Data.SqlClient
Imports PttPosibleApi
Imports Newtonsoft.Json

Module ReceiveOrderFromPTTModule

    Friend Function AddDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopID As Integer, ByVal materialID As Integer,
                                ByVal addAmount As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal, ByVal discountAmount As Decimal,
                                ByVal discountPercent As Decimal, ByVal materialVATType As Integer, ByVal remark As String, ByVal Api60F As String, ByRef resultText As String) As Boolean

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

        Dim dtUnitSmall As New DataTable
        Dim selUnitSmallName As String = ""
        dtUnitSmall = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn, materialID)

        If dtMaterialUnit.Rows.Count = 0 Then
            resultText = "ไม่พบวัตถุดิบที่เลือก"
            Return False
        End If
        selUnitID = dtMaterialUnit.Rows(0)("SelectUnitID")
        selUnitName = dtMaterialUnit.Rows(0)("UnitLargeName")
        selUnitSmallID = dtMaterialUnit.Rows(0)("UnitSmallID")
        selUnitSmallName = dtUnitSmall.Rows(0)("UnitSmallName")

        If Not IsDBNull(dtMaterialUnit.Rows(0)("PTTCode")) Then
            selMaterialCode = dtMaterialUnit.Rows(0)("PTTCode")
        Else
            If Not IsDBNull(dtMaterialUnit.Rows(0)("MaterialCode")) Then
                selMaterialCode = dtMaterialUnit.Rows(0)("MaterialCode")
            End If
        End If
        If Not IsDBNull(dtMaterialUnit.Rows(0)("PTTName")) Then
            selMaterialName = dtMaterialUnit.Rows(0)("PTTName")
        Else
            If Not IsDBNull(dtMaterialUnit.Rows(0)("MaterialName")) Then
                selMaterialName = dtMaterialUnit.Rows(0)("MaterialName")
            End If
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
            DocumentPTTSQL.InsertDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, selDocDetailID, materialID, addAmount,
                                             FormatDecimal(discountPercent, digitDecimal), FormatDecimal(discountAmount, digitDecimal), pricePerUnit, FormatDecimal(selTax, digitDecimal), materialVATType, selUnitSmallID, selUnitID, selUnitName,
                                             selUnitSmallAmount, FormatDecimal(selMaterialNetPrice, digitDecimal), selMaterialCode, selMaterialName, selMaterialSupplierCode, selMaterialSupplierName, remark, Api60F, selUnitSmallName)
            DocumentSQL.UpdateDocSummaryIntoDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function AddDocDetailTank(ByVal globalVariable As GlobalVariable, ByVal docDetailId As Integer, ByVal documentId As Integer, ByVal documentShopID As Integer,
                                     ByVal tankId As Integer, ByVal addAmount As Decimal, ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        Dim dtTank As New DataTable
        Dim expression As String = ""
        Dim foundRows() As DataRow
        Dim tankNo, tankName As String
        Try
            dtTank = DocumentPTTSQL.GetTankDetail(globalVariable.DocDBUtil, globalVariable.DocConn)
            expression = "tankid=" & tankId
            foundRows = dtTank.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                If Not IsDBNull(foundRows(0)("tank_number")) Then
                    tankNo = foundRows(0)("tank_number")
                End If
                If Not IsDBNull(foundRows(0)("tank_name")) Then
                    tankName = foundRows(0)("tank_name")
                End If
            End If
        Catch ex As Exception

        End Try
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentPTTSQL.InsertDocumentDetailTank(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, docDetailId, tankId, tankNo, tankName, addAmount)
            DocumentPTTSQL.SumReceiveOilToDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function AutoAddDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopID As Integer, ByRef resultText As String) As Boolean

        Dim dtTank As New DataTable
        Dim dtMaterialOil As New DataTable
        Dim dtDocDetail As New DataTable
        Dim materialId, unitId, materialTaxType, materialIDRef As Integer
        Dim expression As String = ""
        Dim foundRows() As DataRow
        Dim tankNo, tankName As String
        Dim tankId As Integer
        Dim remark As String = ""
        Dim dbTrans As SqlTransaction
        Try
            dtTank = DocumentPTTSQL.GetTankDetail(globalVariable.DocDBUtil, globalVariable.DocConn)
        Catch ex As Exception

        End Try
        dtMaterialOil = DocumentPTTSQL.GetMaterialOil(globalVariable.DocDBUtil, globalVariable.DocConn)

        If dtMaterialOil.Rows.Count > 0 Then
            For i As Integer = 0 To dtMaterialOil.Rows.Count - 1
                If dtTank.Rows.Count > 0 Then
                    expression = "Grade_ID=" & dtMaterialOil.Rows(i)("MaterialIDRef")
                    foundRows = dtTank.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        If Not IsDBNull(foundRows(0)("tank_number")) Then
                            tankNo = foundRows(0)("tank_number")
                        End If
                        If Not IsDBNull(foundRows(0)("tank_name")) Then
                            tankName = foundRows(0)("tank_name")
                        End If
                    End If
                Else
                    tankNo = ""
                    tankName = ""
                End If

                materialId = dtMaterialOil.Rows(i)("MaterialId")
                materialTaxType = dtMaterialOil.Rows(i)("MaterialTaxtype")
                unitId = dtMaterialOil.Rows(i)("UnitSmallId")
                If AddDocDetail(globalVariable, documentId, documentShopID, materialId, 0, unitId, 0, 0, 0, materialTaxType, remark, "", resultText) = False Then
                    Return False
                End If
            Next
            dtDocDetail = DocumentSQL.GetDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopID)
            dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
            Try
                If dtDocDetail.Rows.Count > 0 Then
                    For i As Integer = 0 To dtDocDetail.Rows.Count - 1
                        If dtTank.Rows.Count > 0 Then
                            expression = "MaterialId=" & dtDocDetail.Rows(i)("ProductID")
                            foundRows = dtMaterialOil.Select(expression)
                            materialIDRef = foundRows(0)("materialIDRef")
                            tankId = 0
                            tankNo = ""
                            tankName = ""
                            expression = "Grade_ID=" & materialIDRef
                            foundRows = dtTank.Select(expression)
                            If foundRows.GetUpperBound(0) >= 0 Then
                                For n As Integer = 0 To foundRows.Length - 1
                                    tankId = foundRows(n)("tank_number")
                                    If Not IsDBNull(foundRows(n)("tank_number")) Then
                                        tankNo = foundRows(n)("tank_number")
                                    End If
                                    If Not IsDBNull(foundRows(n)("tank_name")) Then
                                        tankName = foundRows(n)("tank_name")
                                    End If
                                    If tankId <> 0 Then
                                        DocumentPTTSQL.InsertDocumentDetailTank(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, dtDocDetail.Rows(i)("DocDetailId"), tankId, tankNo, tankName, 0)
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
                dbTrans.Commit()
            Catch ex As Exception
                dbTrans.Rollback()
            End Try

        End If
        Return True
    End Function

    Friend Function ApproveDocument(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer, ByRef resultText As String) As Boolean
        Dim dbTrans As SqlTransaction
        Dim updateDate As DateTime
        Dim strUpdateDate As String
        Dim vendorId As Integer = 0
        Dim dtResult As New DataTable
        Dim documentTypeId As Integer = 0

        dtResult = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId, globalVariable.DocLangID)
        If dtResult.Rows.Count > 0 Then
            documentTypeId = dtResult.Rows(0)("DocumentTypeId")
        End If
        updateDate = Now
        strUpdateDate = FormatDateTime(updateDate)
        dbTrans = globalVariable.DocConn.BeginTransaction
        Try
            DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, globalVariable.StaffID, strUpdateDate)
            DocumentModule.UpdateMaterialDefaultPrice(globalVariable, dbTrans, documentId, documentShopId, vendorId)
            If documentTypeId = globalVariable.DOCUMENTTYPE_DIRECTROPTTNONOIL Then
                MaterialSQL.AutoAddDailyStockMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
                MaterialSQL.AutoAddMonthlyStockMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
                MaterialSQL.AutoAddWeeklyStockMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
            End If
            dbTrans.Commit()
        Catch e1 As Exception
            resultText = e1.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, "ApproveDocument", "ApproveDocument", "99", e1.ToString)
            Return False
        End Try
        If CallPttApi(globalVariable, documentId, documentShopId, globalVariable.DOCUMENTSTATUS_APPROVE, resultText) = False Then
            Try
                DocumentSQL.ReSetDocumentStatus(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId, globalVariable.DOCUMENTSTATUS_WORKING)
            Catch ex As Exception
                resultText = ex.ToString
                Return False
            End Try
            Return False
        End If
        resultText = ""
        Return True
    End Function

    Friend Function CallPttApi(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer, ByVal actionMode As Integer, ByRef resultText As String) As Boolean

        Dim data As New DocumentPTT_Data
        Dim pttErrCode As Integer
        Dim pttErrMsg As String = ""
        Dim actionName As String = ""
        Dim functionName As String = ""
        Dim param As String = ""
        Try

            If DocumentModule.LoadDocument(globalVariable, documentId, documentShopId, data, resultText) = True Then
                Dim pttData As New PttPosibleApi.Model.TSGR
                Dim strArr() As String

                pttData.OBJ_ID = data.DocumentNumber
                pttData.SHIPMENT = data.DeliveryNo
                pttData.ORDERNO = data.InvoiceNo
                pttData.DOCNO = data.ShipmentNo
                pttData.SHIPTO = data.PlantID
                pttData.VAHICLE_ID = data.CarrierName

                Dim documentDate As Date
                strArr = data.DocumentDate.Split("-")
                documentDate = New Date(strArr(0), strArr(1), strArr(2))

                pttData.RECEIVE_DATE = documentDate
                pttData.TOTAL = data.DocSummary.SubTotal
                pttData.TAX = data.DocSummary.TotalVAT
                pttData.GRANDTOTAL = data.DocSummary.GrandTotal

                Dim TaxInvoiceDate As Date
                If Not data.TaxInvoiceDate = Nothing Then
                    strArr = data.TaxInvoiceDate.Split("-")
                    TaxInvoiceDate = New Date(strArr(0), strArr(1), strArr(2))
                    pttData.DUEDATE = TaxInvoiceDate
                End If

                pttData.IS_PAY = 0
                pttData.IS_GI = 0
                pttData.SHIFT_ID = data.ShiftNo
                pttData.DAY_ID = data.ShiftDay
                pttData.MODBY = data.UpdateStaffName
                pttData.TAX_INVOICE_NO = data.TaxInvoiceNo

                Dim InvoiceDate As Date
                If Not data.InvoiceDate = Nothing Then
                    strArr = data.InvoiceDate.Split("-")
                    InvoiceDate = New Date(strArr(0), strArr(1), strArr(2))
                    pttData.INVOICE_NO_DATE = InvoiceDate
                End If
                pttData.INVOICE_NO = data.InvoiceNo

                Dim SaleOrderDate As Date
                If Not data.SaleOrderDate = Nothing Then
                    strArr = data.SaleOrderDate.Split("-")
                    SaleOrderDate = New Date(strArr(0), strArr(1), strArr(2))
                    pttData.SALES_ORDER_NO_DATE = SaleOrderDate
                End If
                pttData.SALES_ORDER_NO = data.SaleOrderNo
                pttData.BUSSINESS_PLACE = data.BusinessID
                pttData.CONTRACT_NO = data.ContractNo

                Dim PurchaseOrderDate As Date
                If Not data.PurchaseOrderDate = Nothing Then
                    strArr = data.PurchaseOrderDate.Split("-")
                    PurchaseOrderDate = New Date(strArr(0), strArr(1), strArr(2))
                    pttData.PURCHASE_ORDER_NO_DATE = PurchaseOrderDate
                End If

                Dim DeliveryDate As Date
                If Not data.DeliveryDate = Nothing Then
                    strArr = data.DeliveryDate.Split("-")
                    DeliveryDate = New Date(strArr(0), strArr(1), strArr(2))
                    pttData.PURCHASE_ORDER_NO_DATE = DeliveryDate
                End If

                pttData.CREATE_BY = data.InsertStaffName
                pttData.GS_NO = data.GS_No
                pttData.INCOTERM = data.Incoterm

                Dim ShipmentDate As Date
                If Not data.ShipmentDate = Nothing Then
                    strArr = data.ShipmentDate.Split("-")
                    ShipmentDate = New Date(strArr(0), strArr(1), strArr(2))
                    pttData.SHIPMENT_DATE = ShipmentDate
                End If

                pttData.SHIPPING_CONDITION = data.ShippingCondition
                pttData.CARRIER_BY = data.CarrierBy
                pttData.DRIVER_NAME = data.DriverName
                pttData.SAEL_NO = data.SealNo
                pttData.TRIP_NO = data.TripNo
                pttData.TOTALAFTERTAX = data.DocSummary.NetPrice
                pttData.DIFFERENCE_FX = 0
                pttData.TRANSFER_TAXCLASS = data.DeliveryCostTaxType
                pttData.TRANSFER_TOTAL = data.DeliveryCostNetPrice
                pttData.TRANSFER_VAT = data.DeliveryCostVAT
                pttData.TRANSFER_AMOUNT = data.DeliveryCost
                pttData.REC_NO = data.DocumentNumber
                pttData.CUS_DOC_NO = data.CustomerDocNo
                pttData.CONTRACT_NO = data.Contractor
                pttData.VAHICLE_ID = data.CarrierNo
                If data.DocumentType = globalVariable.DOCUMENTTYPE_DIRECTROPTT Then
                    pttData.PRODUCT_TYPE = 1
                Else
                    pttData.PRODUCT_TYPE = 0
                End If

                Dim dd As New PttPosibleApi.Model.TSGR_DETAIL
                Dim dtDetail As New List(Of PttPosibleApi.Model.TSGR_DETAIL)
                If data.DocDetailList.Count > 0 Then
                    For i As Integer = 0 To data.DocDetailList.Count - 1
                        If data.DocDetailList(i).Amount > 0 Then
                            dd = New PttPosibleApi.Model.TSGR_DETAIL
                            dd.OBJ_ID = data.DocumentNumber
                            dd.SHIPMENT = data.DeliveryNo
                            dd.ORDERNO = data.InvoiceNo
                            dd.DOCNO = data.ShipmentNo
                            dd.ITEMNO = data.DocDetailList(i).DocDetailID
                            dd.PRICE_LOT = data.DocDetailList(i).PricePerUnit
                            dd.PRICE = data.DocDetailList(i).PricePerUnit
                            'For Receive Stock
                            dd.MAT_ID_LOT = data.DocDetailList(i).MaterialCode
                            dd.QTY_LOT = data.DocDetailList(i).Amount
                            dd.UOM_LOT = data.DocDetailList(i).UnitName
                            'For Sale
                            If data.DocDetailList(i).MaterialCode1 = "" Then
                                dd.MAT_ID = data.DocDetailList(i).MaterialCode
                            Else
                                dd.MAT_ID = data.DocDetailList(i).MaterialCode1
                            End If
                            dd.QTY = data.DocDetailList(i).UnitSmallAmount
                            If data.DocDetailList(i).UnitSmallName = "" Then
                                dd.UOM = data.DocDetailList(i).UnitName
                            Else
                                dd.UOM = data.DocDetailList(i).UnitSmallName
                            End If
                            dd.TAXCLASS = data.DocDetailList(i).MaterialVATType
                            dd.TOTAL = data.DocDetailList(i).MaterialNetPrice
                            dd.TAX = data.DocDetailList(i).MaterialVAT
                            dd.MAT_TEMP = data.DocDetailList(i).MatTemp
                            dd.TEST_TEMP = data.DocDetailList(i).TestTemp
                            dd.TEST_API = data.DocDetailList(i).TestAPI
                            If data.DocumentType = globalVariable.DOCUMENTTYPE_DIRECTROPTT Then
                                dd.QTY_RECIEVE = data.DocDetailList(i).ROSmallAmount
                            Else
                                dd.QTY_RECIEVE = data.DocDetailList(i).UnitSmallAmount
                            End If
                            dtDetail.Add(dd)
                        End If
                    Next
                End If
                pttData.TSGR_DETAIL = dtDetail
                If data.DocumentType = globalVariable.DOCUMENTTYPE_DIRECTROPTT Then
                    Dim dt As New PttPosibleApi.Model.TSGR_TANK_DETAIL
                    Dim tankData As New List(Of PttPosibleApi.Model.TSGR_TANK_DETAIL)
                    dt.OBJ_ID = data.DocumentNumber
                    dt.SHIPMENT = data.DeliveryNo
                    dt.ORDERNO = data.InvoiceNo
                    dt.DOCNO = data.ShipmentNo
                    If data.DocDetailTankList.Count > 0 Then
                        For i As Integer = 0 To data.DocDetailTankList.Count - 1
                            If data.DocDetailTankList(i).ProductAmount > 0 Then
                                dt = New PttPosibleApi.Model.TSGR_TANK_DETAIL
                                dt.ITEMNO = data.DocDetailTankList(i).DocDetailID
                                dt.TANK_ID = data.DocDetailTankList(i).TankID
                                dt.REV_VOLUME = data.DocDetailTankList(i).ProductAmount
                                dt.GUAGE_REV_VOLUME = 0
                                tankData.Add(dt)
                            End If
                        Next
                    End If
                    pttData.TSGR_TANK_DETAIL = tankData

                    functionName = "PttPosibleApi.Api.API_RECEIVE_STOCK_LIGHTOIL"
                    Select Case actionMode
                        Case Is = 2
                            pttData.IS_ACTIVE = True
                            actionName = "PttPosibleApi.Data.Mode.A"
                            Dim pttApi As New PttPosibleApi.Api.API_RECEIVE_STOCK_LIGHTOIL(PttPosibleApi.Data.Mode.A, pttData, pttErrCode, pttErrMsg)
                            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "PttPosibleApi.Api", functionName & " : " & actionName, "77", JsonConvert.SerializeObject(pttData, Formatting.Indented))
                        Case Is = 99
                            pttData.IS_ACTIVE = False
                            pttData.CREATEDATE = documentDate
                            actionName = "PttPosibleApi.Data.Mode.D"
                            Dim pttApi As New PttPosibleApi.Api.API_RECEIVE_STOCK_LIGHTOIL(PttPosibleApi.Data.Mode.D, pttData, pttErrCode, pttErrMsg)
                            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "PttPosibleApi.Api", functionName & " : " & actionName, "77", JsonConvert.SerializeObject(pttData, Formatting.Indented))
                    End Select
                Else
                    functionName = "PttPosibleApi.Api.API_RECEIVE_STOCK_NONOIL"
                    Select Case actionMode
                        Case Is = 2
                            pttData.IS_ACTIVE = True
                            actionName = "PttPosibleApi.Data.Mode.A"
                            Dim pttApi As New PttPosibleApi.Api.API_RECEIVE_STOCK_NONOIL(PttPosibleApi.Data.Mode.A, pttData, pttErrCode, pttErrMsg)
                            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "PttPosibleApi.Api", functionName & " : " & actionName, "77", JsonConvert.SerializeObject(pttData, Formatting.Indented))
                        Case Is = 99
                            pttData.IS_ACTIVE = False
                            pttData.CREATEDATE = documentDate
                            actionName = "PttPosibleApi.Data.Mode.D"
                            Dim pttApi As New PttPosibleApi.Api.API_RECEIVE_STOCK_NONOIL(PttPosibleApi.Data.Mode.D, pttData, pttErrCode, pttErrMsg)
                            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "PttPosibleApi.Api", functionName & " : " & actionName, "77", JsonConvert.SerializeObject(pttData, Formatting.Indented))
                    End Select
                End If

            End If
            If pttErrCode < 0 Then
                DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "PttPosibleApi.Api", functionName & " : " & actionName, "88", pttErrMsg)
                resultText = functionName & " : " & actionName & " Error : " & pttErrMsg
                Return False
            Else
                DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "PttPosibleApi.Api", functionName & " : " & actionName, "22", "Call Api Successfully.")
                resultText = ""
                Return True
            End If
        Catch ex As Exception
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "PttPosibleApi.Api", functionName & " : " & actionName, "88", ex.ToString)
            resultText = functionName & " : " & actionName & " Error : " & Utilitys.Utilitys.ReplaceSuitableStringForSQL(ex.ToString)
            Return False
        End Try

    End Function

    Friend Function CancelDocument(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer, ByRef resultText As String) As Boolean


        Dim dbTrans As SqlTransaction
        Dim strUpdateDate As String
        Dim dtDocument As New DataTable
        dtDocument = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId, globalVariable.DocLangID)
        If dtDocument.Rows.Count = 0 Then
            resultText = globalVariable.MESSAGE_DATANOTFOUND
            Return False
        End If

        'If CheckValidDocumentForCancelDocument(globalVariable, dtDocument.Rows(0)("documentStatus"), dtDocument.Rows(0)("ShopId"), dtDocument.Rows(0)("DocumentDate"), resultText) = False Then
        '    Return False
        'End If

        strUpdateDate = FormatDateTime(Now)
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentSQL.CancelDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, globalVariable.StaffID, strUpdateDate)
            dbTrans.Commit()
        Catch e1 As Exception
            resultText = e1.Message
            dbTrans.Rollback()
            Return False
        End Try
        If dtDocument.Rows(0)("DocumentStatus") = 2 Then
            If CallPttApi(globalVariable, documentId, documentShopId, globalVariable.DOCUMENTSTATUS_CANCEL, resultText) = False Then
                DocumentSQL.ReSetDocumentStatus(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId, dtDocument.Rows(0)("DocumentStatus"))
                Return False
            End If
        End If
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
        Dim strDocumentType As String = globalVariable.DOCUMENTTYPE_DIRECTROPTT & "," & globalVariable.DOCUMENTTYPE_DIRECTROPTTNONOIL
        Try
            dtResult = DocumentSQL.SearchDocument(globalVariable.DocDBUtil, globalVariable.DocConn, strDocumentType, strFromDate, strToDate,
                                                       documentStatus, searchInventoryID, vendorID, vendorGroupID, globalVariable.DocLangID)
            dtDocStatus = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
            If dtResult.Rows.Count > 0 Then
                docList = DocumentModule.InsertResultDataIntoList(globalVariable, globalVariable.DOCUMENTTYPE_DIRECTROPTT, dtResult, dtDocStatus)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function SearchDocument(ByVal globalVariable As GlobalVariable, ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date,
                                   ByVal searchInventoryID As Integer, ByVal vendorID As Integer, ByVal vendorGroupID As Integer, ByVal documentTypeId As Integer,
                                   ByVal taxInvoiceNo As String, ByVal fromTaxInvoiceDate As Date, ByVal toTaxInvoiceDate As Date, ByVal businessPlace As Integer,
                                   ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean

        Dim dtResult, dtDocStatus As DataTable
        Dim strFromDate, strToDate As String
        Dim strFromInvoiceDate, strToInvoiceDate, strDocumentTeype As String
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

        If fromTaxInvoiceDate = Date.MinValue Then
            strFromInvoiceDate = ""
        Else
            strFromInvoiceDate = FormatDate(fromTaxInvoiceDate)
        End If
        If toTaxInvoiceDate = Date.MinValue Then
            strToInvoiceDate = ""
        Else
            strToInvoiceDate = FormatDate(toTaxInvoiceDate)
        End If

        If documentTypeId = -1 Then
            strDocumentTeype = "40,41"
        Else
            strDocumentTeype = documentTypeId
        End If
        Try
            dtResult = DocumentSQL.SearchDocumentPTT(globalVariable.DocDBUtil, globalVariable.DocConn, strDocumentTeype, strFromDate, strToDate, documentStatus, searchInventoryID, vendorID, vendorGroupID, globalVariable.DocLangID, businessPlace, taxInvoiceNo, strFromInvoiceDate, strToInvoiceDate)
            dtDocStatus = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
            If dtResult.Rows.Count > 0 Then
                docList = DocumentModule.InsertResultDataIntoList(globalVariable, globalVariable.DOCUMENTTYPE_DIRECTROPTT, dtResult, dtDocStatus)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function SaveDocumentDataIntoDB(ByVal globalVariable As GlobalVariable, ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date,
    ByVal customerDocNo As String, ByVal customerCode As String, ByVal customerName As String,
    ByVal customerAddress As String, ByVal customerShipTo As String, ByVal customerBillTo As String, ByVal taxInvoiceNo As String, ByVal taxInvoiceDate As Date, ByVal invoiceNo As String,
    ByVal invoiceDate As Date, ByVal saleOrderNo As String, ByVal saleOrderDate As Date, ByVal purchaseOrderNo As String, ByVal purchaseOrderDate As Date, ByVal deliveryNo As String,
    ByVal deliveryDate As Date, ByVal businessId As Integer, ByVal plantId As Integer, ByVal GS_No As String, ByVal paymentDate As Date, ByVal incoterm As String, ByVal contractNo As String,
    ByVal shipmentDate As Date, ByVal shippingCondition As String, ByVal carrierBy As String, ByVal carrierName As String, ByVal driverName As String, ByVal sealNo As String, ByVal tripNo As String,
    ByVal contractor As String, ByVal shippingCost As Decimal, ByVal shippingCostTaxType As Integer, ByVal shipmentNo As String, ByVal shiftID As Integer, ByVal shiftDay As Integer, ByVal shiftNo As Integer,
    ByVal createBy As String, ByRef resultText As String) As Boolean

        Dim strDocDate, strUpdateDate, strTaxInvoiceDate As String
        Dim strInvoiceDate, strSaleOrderDate, strPurchaseOrderDate As String
        Dim strDeliveryDate, strPaymentDate, strShipmentDate As String
        Dim updateDate As DateTime
        Dim newSend As Integer = 0
        Dim dbTrans As SqlTransaction
        Dim transferTax, transferTotal, transferNetPrice As Decimal
        Dim dtProperty As New DataTable
        Dim digitDecimal As Integer = 2
        dtProperty = InventorySQL.GetProperty(globalVariable.DocDBUtil, globalVariable.DocConn)
        If dtProperty.Rows.Count > 0 Then
            digitDecimal = dtProperty.Rows(0)("DigitForRoundingDecimal")
        End If

        strDocDate = FormatDate(documentDate)
        If taxInvoiceDate <> Date.MinValue Then
            strTaxInvoiceDate = FormatDate(taxInvoiceDate)
        Else
            strTaxInvoiceDate = "NULL"
        End If
        If invoiceDate <> Date.MinValue Then
            strInvoiceDate = FormatDate(invoiceDate)
        Else
            strInvoiceDate = "NULL"
        End If
        If saleOrderDate <> Date.MinValue Then
            strSaleOrderDate = FormatDate(saleOrderDate)
        Else
            strSaleOrderDate = "NULL"
        End If
        If purchaseOrderDate <> Date.MinValue Then
            strPurchaseOrderDate = FormatDate(purchaseOrderDate)
        Else
            strPurchaseOrderDate = "NULL"
        End If
        If deliveryDate <> Date.MinValue Then
            strDeliveryDate = FormatDate(deliveryDate)
        Else
            strDeliveryDate = "NULL"
        End If
        If paymentDate <> Date.MinValue Then
            strPaymentDate = FormatDate(paymentDate)
        Else
            strPaymentDate = "NULL"
        End If
        If shipmentDate <> Date.MinValue Then
            strShipmentDate = FormatDate(shipmentDate)
        Else
            strShipmentDate = "NULL"
        End If

        updateDate = Now
        strUpdateDate = FormatDateTime(updateDate)
        newSend = 0
        transferTotal = shippingCost

        Select Case shippingCostTaxType
            Case globalVariable.TAXTYPE_NOVAT
                transferTax = 0
            Case globalVariable.TAXTYPE_EXCLUDEVAT
                transferTax = (shippingCost * globalVariable.DefaultShopVAT) / 100
                transferNetPrice = transferTotal
            Case globalVariable.TAXTYPE_INCLUDEVAT
                transferTax = (shippingCost * globalVariable.DefaultShopVAT) / (100 + globalVariable.DefaultShopVAT)
                transferNetPrice = (transferTotal - transferTax)
        End Select
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentPTTSQL.UpdateDocumentFormPTT(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentID, documentShopID, inventoryID, strDocDate, customerDocNo, customerCode, customerName, customerAddress, customerShipTo, customerBillTo,
                                              taxInvoiceNo, strTaxInvoiceDate, invoiceNo, strInvoiceDate, saleOrderNo, strSaleOrderDate, purchaseOrderNo, strPurchaseOrderDate, deliveryNo, strDeliveryDate, businessId, plantId, strPaymentDate, incoterm, contractNo,
                                              strShipmentDate, shippingCondition, carrierBy, carrierName, driverName, sealNo, tripNo, contractor, shippingCostTaxType, FormatDecimal(transferTotal, globalVariable.DigitForRoundingDecimal),
                                              FormatDecimal(transferTax, globalVariable.DigitForRoundingDecimal), FormatDecimal(transferNetPrice, globalVariable.DigitForRoundingDecimal), shipmentNo, shiftID, shiftDay, shiftNo, GS_No, strUpdateDate, globalVariable.StaffID, createBy)
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "SaveDocumentDataIntoDB", "SaveDocumentDataIntoDB", "99", ex.ToString)
            Return False
        End Try

        resultText = ""
        Return True
    End Function

    Friend Function UpdateDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer,
                                    ByVal materialID As Integer, ByVal addAmount As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal,
                                    ByVal discountAmount As Decimal, ByVal discountPercent As Decimal, ByVal materialVATType As Integer, ByVal remark As String, ByVal Api60F As String,
                                    ByVal matTemp As Decimal, ByVal testTemp As Decimal, ByVal testApi As Decimal, ByRef resultText As String) As Boolean

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
        Dim dtUnitSmall As New DataTable
        Dim selUnitSmallName As String = ""
        dtUnitSmall = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn, materialID)

        selUnitID = dtMaterialUnit.Rows(0)("SelectUnitID")
        selUnitName = dtMaterialUnit.Rows(0)("UnitLargeName")
        selUnitSmallID = dtMaterialUnit.Rows(0)("UnitSmallID")
        selUnitSmallName = dtUnitSmall.Rows(0)("UnitSmallName")

        If Not IsDBNull(dtMaterialUnit.Rows(0)("PTTCode")) Then
            selMaterialCode = dtMaterialUnit.Rows(0)("PTTCode")
        Else
            If Not IsDBNull(dtMaterialUnit.Rows(0)("MaterialCode")) Then
                selMaterialCode = dtMaterialUnit.Rows(0)("MaterialCode")
            End If
        End If
        If Not IsDBNull(dtMaterialUnit.Rows(0)("PTTName")) Then
            selMaterialName = dtMaterialUnit.Rows(0)("PTTName")
        Else
            If Not IsDBNull(dtMaterialUnit.Rows(0)("MaterialName")) Then
                selMaterialName = dtMaterialUnit.Rows(0)("MaterialName")
            End If
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
            DocumentModule.CalculateDocDetailAllPrice(globalVariable, addAmount, pricePerUnit, discountPercent, discountAmount, materialVATType,
                                                      selTotalPriceBeforeDiscount, selDiscountPrice, selTax, selMaterialNetPrice)
            DocumentPTTSQL.UpdateDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, docDetailId, materialID, addAmount,
                                             discountPercent, discountAmount, pricePerUnit, selTax, materialVATType, selUnitSmallID, selUnitID, selUnitName,
                                             selUnitSmallAmount, selMaterialNetPrice, selMaterialCode, selMaterialName, selMaterialSupplierCode, selMaterialSupplierName, remark, Api60F,
                                             matTemp, testTemp, testApi, selUnitSmallName)
            DocumentSQL.UpdateDocSummaryIntoDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function UpdateDocDetailTank(ByVal globalVariable As GlobalVariable, ByVal docDetailId As Integer, ByVal documentId As Integer, ByVal documentShopID As Integer,
                                        ByVal tankId As Integer, ByVal addAmount As Decimal, ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentPTTSQL.UpdateDocumentDetailTank(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, docDetailId, tankId, addAmount)
            DocumentPTTSQL.SumReceiveOilToDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
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
                DocumentPTTSQL.DeleteDocumentDetailTank(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, docDetailId(i))
                DocumentPTTSQL.SumReceiveOilToDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
            Next i
            DocumentSQL.UpdateDocSummaryIntoDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
            dbTrans.Commit()
        Catch e1 As Exception
            resultText = e1.ToString
            dbTrans.Rollback()
            Return False
        End Try

        resultText = ""
        Return True
    End Function

    Friend Function DeleteDocDetailTank(ByVal globalVariable As GlobalVariable, ByVal docDetailId As Integer, ByVal documentId As Integer, ByVal documentShopID As Integer,
                                        ByVal tankId As Integer, ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentPTTSQL.DeleteDocumentDetailTank(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, docDetailId, tankId)
            DocumentPTTSQL.SumReceiveOilToDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function LoadDocumentDetailTank(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer,
                                           ByRef docData As List(Of DocumentDetailTank_Data), ByRef resultText As String) As Boolean

        Dim dtResult As New DataTable
        Try
            dtResult = DocumentPTTSQL.GetDocumentDetailTank(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId)
            If dtResult.Rows.Count > 0 Then
                For i As Integer = 0 To dtResult.Rows.Count - 1
                    docData.Add(DocumentDetailTank_Data.NewTankData(dtResult.Rows(i)("DocDetailId"), dtResult.Rows(i)("DocumentId"), dtResult.Rows(i)("ShopId"), dtResult.Rows(i)("TankId"), dtResult.Rows(i)("TankNo"), dtResult.Rows(i)("TankName"), dtResult.Rows(i)("ProductAmount")))
                Next
            End If
        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function LoadDocumentDetailTank(ByVal globalVariable As GlobalVariable, ByVal docDetailId As Integer, ByVal documentId As Integer, ByVal documentShopId As Integer,
                                           ByRef docData As List(Of DocumentDetailTank_Data), ByRef resultText As String) As Boolean

        Dim dtResult As New DataTable
        Try
            dtResult = DocumentPTTSQL.GetDocumentDetailTank(globalVariable.DocDBUtil, globalVariable.DocConn, docDetailId, documentId, documentShopId)
            If dtResult.Rows.Count > 0 Then
                For i As Integer = 0 To dtResult.Rows.Count - 1
                    docData.Add(DocumentDetailTank_Data.NewTankData(dtResult.Rows(i)("DocDetailId"), dtResult.Rows(i)("DocumentId"), dtResult.Rows(i)("ShopId"), dtResult.Rows(i)("TankId"), dtResult.Rows(i)("TankNo"), dtResult.Rows(i)("TankName"), dtResult.Rows(i)("ProductAmount")))
                Next
            End If
        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

End Module
