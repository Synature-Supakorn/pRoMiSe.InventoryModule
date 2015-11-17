Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Public Class ReceiveOrderFromPTTController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByRef docData As DocumentPTT_Data, ByRef resultText As String) As Boolean
        If ReceiveOrderFromPTTModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date,
    ByVal customerDocNo As String, ByVal customerCode As String, ByVal customerName As String, ByVal customerAddress As String, ByVal customerShipTo As String,
    ByVal customerBillTo As String, ByVal taxInvoiceNo As String, ByVal taxInvoiceDate As Date, ByVal invoiceNo As String,
    ByVal invoiceDate As Date, ByVal saleOrderNo As String, ByVal saleOrderDate As Date, ByVal purchaseOrderNo As String, ByVal purchaseOrderDate As Date, ByVal deliveryNo As String,
    ByVal deliveryDate As Date, ByVal businessId As Integer, ByVal plantId As Integer, ByVal GS_No As String, ByVal paymentDate As Date, ByVal incoterm As String, ByVal contractNo As String,
    ByVal shipmentDate As Date, ByVal shippingCondition As String, ByVal carrierBy As String, ByVal carrierName As String, ByVal driverName As String, ByVal sealNo As String, ByVal tripNo As String,
    ByVal contractor As String, ByVal shippingCost As Decimal, ByVal shippingCostTaxType As Integer, ByVal shipmentNo As String,
    ByVal shiftID As Integer, ByVal createBy As String, ByRef docData As DocumentPTT_Data, ByRef resultText As String) As Boolean

        Dim shiftDay As Integer
        Dim shiftNo As Integer
        Dim dtShift As New DataTable
        dtShift = DocumentSQL.ListShiftData(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryID, documentDate, shiftID)
        If dtShift.Rows.Count > 0 Then
            shiftDay = dtShift.Rows(0)("DAY_ID")
            shiftNo = dtShift.Rows(0)("SHIFT_NO")
        Else
            shiftDay = documentDate.Day
            shiftNo = 1
        End If
        If ReceiveOrderFromPTTModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, globalVariable.DOCUMENTTYPE_DIRECTROPTT, documentDate,
                                                            customerDocNo, customerCode, customerName, customerAddress, customerShipTo, customerBillTo, taxInvoiceNo,
                                                            taxInvoiceDate, invoiceNo, invoiceDate, saleOrderNo, saleOrderDate, purchaseOrderNo, purchaseOrderDate,
                                                            deliveryNo, deliveryDate, businessId, plantId, GS_No, paymentDate, incoterm, contractNo, shipmentDate,
                                                            shippingCondition, carrierBy, carrierName, driverName, sealNo, tripNo, contractor, shippingCost,
                                                            shippingCostTaxType, shipmentNo, shiftID, shiftDay, shiftNo, createBy, resultText) = False Then
            Return False
        End If
        If ReceiveOrderFromPTTModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AddMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As DocumentPTT_Data, ByVal materialID As Integer,
                                                              ByVal materailQty As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal,
                                                              ByVal discountAmount As Decimal, ByVal discountPercent As Decimal, ByVal materialVATType As Integer,
                                                              ByVal remark As String, ByVal Api60F As String, ByRef resultText As String) As Boolean
        If ReceiveOrderFromPTTModule.AddDocDetail(globalVariable, documentID, documentShopID, materialID, materailQty, materialUnitLargeID, pricePerUnit, discountAmount,
                                                 discountPercent, materialVATType, remark, Api60F, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AddMaterialInDocDetailTank(ByVal docDetailId As Integer, ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal tankId As Integer, ByVal addAmount As Decimal,
                                               ByRef docData As List(Of DocumentDetailTank_Data), ByRef resultText As String) As Boolean
        If ReceiveOrderFromPTTModule.AddDocDetailTank(globalVariable, docDetailId, documentID, documentShopID, tankId, addAmount, resultText) = False Then
            Return False
        End If
        Return ReceiveOrderFromPTTModule.LoadDocumentDetailTank(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AutoAddDocDetail(ByVal documentId As Integer, ByVal documentShopID As Integer, ByRef resultText As String) As Boolean
        Return ReceiveOrderFromPTTModule.AutoAddDocDetail(globalVariable, documentId, documentShopID, resultText)
    End Function

    Public Function CreateNewDocument(ByVal documentTypeId As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date, ByVal customerDocNo As String, ByVal customerCode As String,
    ByVal customerName As String, ByVal customerAddress As String, ByVal customerShipTo As String,
    ByVal customerBillTo As String, ByVal taxInvoiceNo As String, ByVal taxInvoiceDate As Date, ByVal invoiceNo As String, ByVal invoiceDate As Date,
    ByVal saleOrderNo As String, ByVal saleOrderDate As Date, ByVal purchaseOrderNo As String, ByVal purchaseOrderDate As Date, ByVal deliveryNo As String,
    ByVal deliveryDate As Date, ByVal businessId As Integer, ByVal plantId As Integer, ByVal GS_No As String, ByVal paymentDate As Date, ByVal incoterm As String,
    ByVal contractNo As String, ByVal shipmentDate As Date, ByVal shippingCondition As String, ByVal carrierBy As String, ByVal carrierName As String, ByVal driverName As String,
    ByVal sealNo As String, ByVal tripNo As String, ByVal contractor As String, ByVal shippingCost As Decimal, ByVal shippingCostTaxType As Integer, ByVal shipmentNo As String,
    ByVal shiftID As Integer, ByVal createBy As String, ByRef docData As DocumentPTT_Data, ByRef resultText As String) As Boolean

        Dim shiftDay As Integer
        Dim shiftNo As Integer
        Dim dtShift As New DataTable
        Dim dtResult As New DataTable

        If documentTypeId = 41 Then
            If CheckValidDocumentDateForSaveNewDocument(globalVariable, inventoryID, documentDate, resultText) = False Then
                Return False
            End If
        End If

        dtShift = DocumentSQL.ListShiftData(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryID, documentDate, shiftID)
        If dtShift.Rows.Count > 0 Then
            shiftDay = dtShift.Rows(0)("DAY_ID")
            shiftNo = dtShift.Rows(0)("SHIFT_NO")
        Else
            shiftDay = documentDate.Day
            shiftNo = 1
        End If
        If DocumentModule.CreateNewDocument(globalVariable, documentTypeId, inventoryID, inventoryID, documentDate, docData, resultText) = False Then
            Return False
        End If
        Select Case documentTypeId
            Case Is = globalVariable.DOCUMENTTYPE_DIRECTROPTT
                If ReceiveOrderFromPTTModule.AutoAddDocDetail(globalVariable, docData.DocumentID, inventoryID, resultText) = False Then
                    Return False
                End If
        End Select

        dtResult = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, docData.DocumentID, docData.DocumentShopID, globalVariable.DocLangID)
        Try

            Dim docNumber As String = GetDocumentHeader(dtResult.Rows(0)("DocumentTypeHeader"), dtResult.Rows(0)("DocumentYear"), dtResult.Rows(0)("DocumentMonth"), dtResult.Rows(0)("DocumentNumber"), globalVariable.DocYearSettingType)

            shipmentNo = "D" & Right(documentDate.Year, 2) & Right(documentDate.Month, 2) & Right((documentDate.Day + 100), 2) & Right(docNumber, 4)
            If deliveryNo = "" Then
                deliveryNo = "S" & Right(documentDate.Year, 2) & Right(documentDate.Month, 2) & Right((documentDate.Day + 100), 2) & Right(docNumber, 4)
            End If
            If invoiceNo = "" Then
                invoiceNo = "O" & Right(documentDate.Year, 2) & Right(documentDate.Month, 2) & Right((documentDate.Day + 100), 2) & Right(docNumber, 4)
            End If
            If customerDocNo = "" Then
                customerDocNo = "R" & Right(documentDate.Year, 2) & Right(documentDate.Month, 2) & Right((documentDate.Day + 100), 2) & Right(docNumber, 4)
            End If

        Catch ex As Exception
            shipmentNo = ""
        End Try
          
        If ReceiveOrderFromPTTModule.SaveDocumentDataIntoDB(globalVariable, docData.DocumentID, inventoryID, documentTypeId, documentDate,
                                                            customerDocNo, customerCode, customerName, customerAddress, customerShipTo, customerBillTo, taxInvoiceNo,
                                                            taxInvoiceDate, invoiceNo, invoiceDate, saleOrderNo, saleOrderDate, purchaseOrderNo, purchaseOrderDate,
                                                            deliveryNo, deliveryDate, businessId, plantId, GS_No, paymentDate, incoterm, contractNo, shipmentDate,
                                                            shippingCondition, carrierBy, carrierName, driverName, sealNo, tripNo, contractor, shippingCost,
                                                            shippingCostTaxType, shipmentNo, shiftID, shiftDay, shiftNo, createBy, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, docData.DocumentID, docData.DocumentShopID, docData, resultText)
    End Function

    Public Function CancelDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As DocumentPTT_Data, ByRef resultText As String) As Boolean
        If ReceiveOrderFromPTTModule.CancelDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function DeleteMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal strDocDetailId As String, ByRef docData As DocumentPTT_Data, ByRef resultText As String) As Boolean
        If ReceiveOrderFromPTTModule.DeleteDocDetail(globalVariable, documentID, documentShopID, strDocDetailId, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function DeleteMaterialInDocDetailTank(ByVal docDetailId As Integer, ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal tankId As Integer, ByVal addAmount As Decimal,
                                                  ByRef docData As List(Of DocumentDetailTank_Data), ByRef resultText As String) As Boolean
        If ReceiveOrderFromPTTModule.DeleteDocDetailTank(globalVariable, docDetailId, documentID, documentShopID, tankId, resultText) = False Then
            Return False
        End If
        Return ReceiveOrderFromPTTModule.LoadDocumentDetailTank(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function ListShippingCondition(ByRef conditionData As List(Of ShippingCondition_Data), ByRef resultText As String) As Boolean
        conditionData = New List(Of ShippingCondition_Data)
        Try
            conditionData = New List(Of ShippingCondition_Data)
            conditionData.Add(ShippingCondition_Data.NewShippingCondition("ลูกค้ารับเอง"))
            conditionData.Add(ShippingCondition_Data.NewShippingCondition("ปตท. ส่งให้"))
        Catch ex As Exception
            resultText = ex.ToString
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function ListShift(ByVal inventoryId As Integer, ByRef shiftdata As List(Of Shift_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtResult As DataTable
        shiftdata = New List(Of Shift_Data)
        Try
            dtResult = DocumentSQL.ListShiftData(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryId)
            shiftdata = New List(Of Shift_Data)
            If dtResult.Rows.Count > 0 Then
                For i = 0 To dtResult.Rows.Count - 1
                    shiftdata.Add(Shift_Data.NewShift(dtResult.Rows(i)("PERIOD_ID"), dtResult.Rows(i)("SHIFT_NO")))
                Next i
            Else
                'shiftdata.Add(Shift_Data.NewShift(1, 1))
                resultText = "วันที่รับสินค้าที่เลือกไม่พบ กะที่ การขาย กรุณาตรวจสอบวันที่รับสินค้าใหม่อีกครั้ง"
            End If
        Catch ex As Exception
            resultText = ex.ToString
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function ListShift(ByVal inventoryId As Integer, ByVal receiveDate As Date, ByRef shiftdata As List(Of Shift_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtResult As DataTable
        shiftdata = New List(Of Shift_Data)
        Try
            dtResult = DocumentSQL.ListShiftData(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryId, receiveDate)
            shiftdata = New List(Of Shift_Data)
            If dtResult.Rows.Count > 0 Then
                For i = 0 To dtResult.Rows.Count - 1
                    shiftdata.Add(Shift_Data.NewShift(dtResult.Rows(i)("PERIOD_ID"), dtResult.Rows(i)("SHIFT_NO")))
                Next i
            Else
                'shiftdata.Add(Shift_Data.NewShift(1, 0))
                resultText = "วันที่รับสินค้าที่เลือกไม่พบ กะที่ การขาย กรุณาตรวจสอบวันที่รับสินค้าใหม่อีกครั้ง"
                Return False
            End If
        Catch ex As Exception
            resultText = ex.ToString
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function ListPlant(ByRef businessPlaceData As List(Of Plant_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtResult As DataTable
        businessPlaceData = New List(Of Plant_Data)
        Try
            dtResult = DocumentSQL.ListPlant(globalVariable.DocDBUtil, globalVariable.DocConn)
            businessPlaceData = New List(Of Plant_Data)
            For i = 0 To dtResult.Rows.Count - 1
                businessPlaceData.Add(Plant_Data.NewPlant(dtResult.Rows(i)("DEPOT_ID"), dtResult.Rows(i)("DEPOT_NAME")))
            Next i
        Catch ex As Exception
            resultText = ex.ToString
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function ListCustomer(ByRef customerData As List(Of Customer_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtResult As DataTable
        customerData = New List(Of Customer_Data)
        Try
            dtResult = DocumentSQL.ListCustomer(globalVariable.DocDBUtil, globalVariable.DocConn)
            customerData = New List(Of Customer_Data)
            For i = 0 To dtResult.Rows.Count - 1
                customerData.Add(Customer_Data.NewCustomerData(dtResult.Rows(i)("DEPOT"), dtResult.Rows(i)("SITENAME"),
                                               dtResult.Rows(i)("SITEADD"), dtResult.Rows(i)("SITEADD"), dtResult.Rows(i)("SITEADD")))
            Next i
        Catch ex As Exception
            resultText = ex.ToString
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function ListBusinessPlace(ByRef businessPlaceData As List(Of BusinessPlace_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtResult As DataTable
        businessPlaceData = New List(Of BusinessPlace_Data)
        Try
            dtResult = DocumentSQL.ListBusinessPlace(globalVariable.DocDBUtil, globalVariable.DocConn)
            businessPlaceData = New List(Of BusinessPlace_Data)
            For i = 0 To dtResult.Rows.Count - 1
                businessPlaceData.Add(BusinessPlace_Data.NewBusinessPlace(dtResult.Rows(i)("BUS_ID"), dtResult.Rows(i)("BUS_NAME")))
            Next i
        Catch ex As Exception
            resultText = ex.ToString
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function LoadDocumentDetailPTT(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal oldDocumentID As Integer,
                                              ByVal oldDocumentShopID As Integer, ByRef documentData As DocumentPTT_Data, ByRef resultText As String) As Boolean
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, documentData, resultText)
    End Function

    Public Function LoadDocDetailTank(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As List(Of DocumentDetailTank_Data),
                                      ByRef resultText As String) As Boolean
        Return ReceiveOrderFromPTTModule.LoadDocumentDetailTank(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function LoadDocDetailTank(ByVal docDetialId As Integer, ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As List(Of DocumentDetailTank_Data),
                                      ByRef resultText As String) As Boolean
        Return ReceiveOrderFromPTTModule.LoadDocumentDetailTank(globalVariable, docDetialId, documentID, documentShopID, docData, resultText)
    End Function

    Public Function SearchDocument(ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date, ByVal inventoryID As Integer,
                                   ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean

        Return ReceiveOrderFromPTTModule.SearchDocument(globalVariable, documentStatus, startDate, endDate, inventoryID, -1, -1, docList, resultText)
    End Function

    Public Function SearchDocument(ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date, ByVal inventoryID As Integer,ByVal documentTypeId As Integer,
                                   ByVal taxInvoiceNo As String, ByVal fromTaxInvoiceDate As Date, ByVal toTaxInvoiceDate As Date, ByVal businessPlace As Integer,
                                   ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean

        Return ReceiveOrderFromPTTModule.SearchDocument(globalVariable, documentStatus, startDate, endDate, inventoryID, -1, -1, documentTypeId, taxInvoiceNo, fromTaxInvoiceDate, toTaxInvoiceDate, businessPlace, docList, resultText)
    End Function

    Public Function SaveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date,
    ByVal customerDocNo As String, ByVal customerCode As String, ByVal customerName As String, ByVal customerAddress As String, ByVal customerShipTo As String,
    ByVal customerBillTo As String, ByVal taxInvoiceNo As String, ByVal taxInvoiceDate As Date, ByVal invoiceNo As String, ByVal invoiceDate As Date,
    ByVal saleOrderNo As String, ByVal saleOrderDate As Date, ByVal purchaseOrderNo As String, ByVal purchaseOrderDate As Date, ByVal deliveryNo As String,
    ByVal deliveryDate As Date, ByVal businessId As Integer, ByVal plantId As Integer, ByVal GS_No As String, ByVal paymentDate As Date, ByVal incoterm As String,
    ByVal contractNo As String, ByVal shipmentDate As Date, ByVal shippingCondition As String, ByVal carrierBy As String, ByVal carrierName As String, ByVal driverName As String, ByVal sealNo As String, ByVal tripNo As String,
    ByVal contractor As String, ByVal shippingCost As Decimal, ByVal shippingCostTaxType As Integer, ByVal shipmentNo As String,
    ByVal shiftID As Integer, ByVal createBy As String, ByRef docData As DocumentPTT_Data, ByRef resultText As String) As Boolean

        Dim shiftDay As Integer
        Dim shiftNo As Integer
        If ReceiveOrderFromPTTModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, globalVariable.DOCUMENTTYPE_DIRECTROPTT, documentDate,
                                                              customerDocNo, customerCode, customerName, customerAddress, customerShipTo, customerBillTo, taxInvoiceNo,
                                                              taxInvoiceDate, invoiceNo, invoiceDate, saleOrderNo, saleOrderDate, purchaseOrderNo, purchaseOrderDate,
                                                              deliveryNo, deliveryDate, businessId, plantId, GS_No, paymentDate, incoterm, contractNo, shipmentDate,
                                                              shippingCondition, carrierBy, carrierName, driverName, sealNo, tripNo, contractor, shippingCost,
                                                              shippingCostTaxType, shipmentNo, shiftID, shiftDay, shiftNo, createBy, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, inventoryID, docData, resultText)
    End Function

    Public Function UpdateMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer, ByVal materialID As Integer,
                                                                 ByVal materailQty As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal, ByVal discountAmount As Decimal,
                                                                 ByVal discountPercent As Decimal, ByVal materialVATType As Integer,
                                                                 ByVal remark As String, ByVal Api60F As String, matTemp As Decimal, ByVal testTemp As Decimal, ByVal testApi As Decimal,
                                                                 ByRef docData As DocumentPTT_Data, ByRef resultText As String) As Boolean
        If ReceiveOrderFromPTTModule.UpdateDocDetail(globalVariable, documentID, documentShopID, docDetailId, materialID, materailQty, materialUnitLargeID, pricePerUnit, discountAmount,
                                                    discountPercent, materialVATType, remark, Api60F, matTemp, testTemp, testApi, resultText) = False Then
            Return False
        End If
        Dim dtTank As New DataTable
        Dim dtPTTTank As New DataTable
        Dim dtMatOil As New DataTable
        Dim objTank As New List(Of DocumentDetailTank_Data)
        Dim expression As String = ""
        Dim foundRows() As DataRow

        dtPTTTank = DocumentPTTSQL.GetTankDetail(globalVariable.DocDBUtil, globalVariable.DocConn)
        dtMatOil = DocumentPTTSQL.GetMaterialOil(globalVariable.DocDBUtil, globalVariable.DocConn)
        expression = "MaterialID=" & materialID
        foundRows = dtMatOil.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            expression = "Grade_ID=" & foundRows(0)("materialIDRef")
            foundRows = dtPTTTank.Select(expression)
            If foundRows.Length < 2 Then
                dtTank = GetDocumentDetailTank(globalVariable.DocDBUtil, globalVariable.DocConn, docDetailId, documentID, documentShopID)
                If dtTank.Rows.Count > 0 Then
                    UpdateMaterialInDocDetailTank(docDetailId, documentID, documentShopID, dtTank.Rows(0)("TankId"), materailQty, objTank, resultText)
                End If
            End If
        End If

        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function UpdateMaterialInDocDetailTank(ByVal docDetailId As Integer, ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal tankId As Integer, ByVal addAmount As Decimal,
                                                  ByRef docData As List(Of DocumentDetailTank_Data), ByRef resultText As String) As Boolean
        If ReceiveOrderFromPTTModule.UpdateDocDetailTank(globalVariable, docDetailId, documentID, documentShopID, tankId, addAmount, resultText) = False Then
            Return False
        End If
        Return ReceiveOrderFromPTTModule.LoadDocumentDetailTank(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

End Class
