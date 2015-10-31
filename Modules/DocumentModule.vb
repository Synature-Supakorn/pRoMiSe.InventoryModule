Imports pRoMiSe.Utilitys.Utilitys
Imports System.Data.SqlClient

Module DocumentModule

    Enum RoundType
        None
        ZeroOne
        Point5
        Point25
        ZeroOneDown
        Point5Down
        Point25Down
        ZeroOneByRoundingFunction
        Point5ByRoundingFunction
        Point25ByRoundingFunction
        ZeroOneByRoundingFunctionHalfRoundDown
        Point5ByRoundingFunctionHalfRoundDown
        Point25ByRoundingFunctionHalfRoundDown
    End Enum

    Friend Function CreateNewDocument(ByVal globalVariable As GlobalVariable, ByVal documentTypeID As Integer, ByVal inventoryID As Integer, ByVal toInventoryID As Integer,
                                      ByVal documentDate As Date, ByRef documentData As Document_Data, ByRef resultText As String) As Boolean

        Dim newDocID As Integer
        Dim strUpdateDate As String
        Dim strDocDate As String
        Dim dbTrans As SqlTransaction
        Dim newDocumentNumber As Integer
        Dim docYear, docMonth As Integer

        strUpdateDate = FormatDateTime(Now)
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try

            newDocID = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, inventoryID)
            If documentDate <> Date.MinValue Then
                strDocDate = FormatDate(documentDate)
            Else
                strDocDate = FormatDate(Date.Now)
            End If

            DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocID, inventoryID, inventoryID, toInventoryID, documentTypeID, globalVariable.DOCUMENTSTATUS_WORKING, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
            newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, inventoryID, documentTypeID, documentDate, docYear, docMonth)
            DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocID, inventoryID, documentTypeID, docMonth, docYear, newDocumentNumber, globalVariable.DocLangID)

            dbTrans.Commit()
            documentData.DocumentID = newDocID
            documentData.DocumentShopID = inventoryID
            documentData.DocumentType = documentTypeID
            Return True
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DocumentModule", "CreateNewDocument", "99", ex.ToString)
            Return False
        End Try
    End Function

    Friend Function CreateNewDocument(ByVal globalVariable As GlobalVariable, ByVal documentTypeID As Integer, ByVal inventoryID As Integer, ByVal toInventoryID As Integer,
                                      ByVal documentDate As Date, ByRef documentData As DocumentPTT_Data, ByRef resultText As String) As Boolean

        Dim newDocID As Integer
        Dim strUpdateDate As String
        Dim strDocDate As String
        Dim dbTrans As SqlTransaction

        Dim newDocumentNumber As Integer
        Dim docYear, docMonth As Integer

        strUpdateDate = FormatDateTime(Now)
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try

            newDocID = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, inventoryID)
            If documentDate <> Date.MinValue Then
                strDocDate = FormatDate(documentDate)
            Else
                strDocDate = "NULL"
            End If
            DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocID, inventoryID, inventoryID, toInventoryID, documentTypeID, globalVariable.DOCUMENTSTATUS_WORKING, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
            newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, inventoryID, documentTypeID, documentDate, docYear, docMonth)
            DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocID, inventoryID, documentTypeID, docMonth, docYear, newDocumentNumber, globalVariable.DocLangID)

            dbTrans.Commit()
            documentData.DocumentID = newDocID
            documentData.DocumentShopID = inventoryID
            documentData.DocumentType = documentTypeID

            Return True
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DocumentModule", "CreateNewDocumentPTT", "99", ex.ToString)
            Return False
        End Try
    End Function

    Friend Function CreateNewDocumentFromReferedDocument(ByVal globalVariable As GlobalVariable, ByVal documentTypeID As Integer, ByVal inventoryID As Integer,
                                                         ByVal documentRefId As Integer, ByVal documentRefShopId As Integer, ByVal documentDate As Date,
                                                         ByVal IncludePriceInform As Boolean, ByRef documentData As Document_Data, ByRef resultText As String) As Boolean
        Dim newDocID As Integer
        Dim strUpdateDate As String
        Dim strDocDate As String
        Dim dbTrans As SqlTransaction
        Dim dtResult As New DataTable
        Dim dtDocRef As New DataTable
        Dim dtMainDoc As New DataTable
        Dim strSelect As String = ""
        Dim rResult() As DataRow
        Dim dclTemp As Decimal
        Dim newDocumentNumber As Integer
        Dim docYear, docMonth As Integer

        documentData = New Document_Data
        dtResult = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentRefId, documentRefShopId, globalVariable.DocLangID)
        If dtResult.Rows.Count > 0 Then
            documentData.DocumentRefNumber = GetDocumentHeader(dtResult.Rows(0)("DocumentTypeHeader"), dtResult.Rows(0)("DocumentYear"), dtResult.Rows(0)("DocumentMonth"), dtResult.Rows(0)("DocumentNumber"), globalVariable.DocYearSettingType)
            documentData.VendorGroupID = dtResult.Rows(0)("VendorGroupID")
            documentData.VendorID = dtResult.Rows(0)("VendorId")
        End If
        strUpdateDate = FormatDate(Now)

        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            newDocID = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, inventoryID)
            If documentDate <> Date.MinValue Then
                strDocDate = FormatDate(documentDate)
            Else
                strDocDate = "NULL"
            End If
            DocumentSQL.CreateNewDocumentFromReferedDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocID, inventoryID, documentTypeID, documentRefId, documentRefShopId, documentData.DocumentNumber, strUpdateDate, strDocDate, globalVariable.StaffID)
            DocumentSQL.CopyDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentRefId, documentRefShopId, newDocID, inventoryID, dtResult.Rows(0)("DocumentStatus"))
            newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, inventoryID, documentTypeID, documentDate, docYear, docMonth)
            DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocID, inventoryID, documentTypeID, docMonth, docYear, newDocumentNumber, globalVariable.DocLangID)

            If IncludePriceInform = True Then
                dtMainDoc = DocumentSQL.GetMaterialFromOriginalDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentRefId, documentRefShopId)
                For i = 0 To dtMainDoc.Rows.Count - 1
                    strSelect &= dtMainDoc.Rows(i)("ProductID") & ", "
                Next i
                If strSelect <> "" Then
                    strSelect = Mid(strSelect, 1, Len(strSelect) - 2)
                Else
                    strSelect = "-1"
                End If

                dtDocRef = DocumentSQL.GetMaterialFromReferenceDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentRefId, documentRefShopId, newDocID, inventoryID, strSelect, documentTypeID)
                For i = 0 To dtMainDoc.Rows.Count - 1
                    'Set UnitRatio
                    If dtMainDoc.Rows(i)("ProductAmount") = 0 Then
                        dtMainDoc.Rows(i)("UnitRatio") = dtMainDoc.Rows(i)("UnitSmallAmount")
                    Else
                        dtMainDoc.Rows(i)("UnitRatio") = dtMainDoc.Rows(i)("UnitSmallAmount") / dtMainDoc.Rows(i)("ProductAmount")
                        If dtMainDoc.Rows(i)("UnitRatio") = 0 Then
                            dtMainDoc.Rows(i)("UnitRatio") = 1
                        End If
                    End If
                    strSelect = "ProductID = " & dtMainDoc.Rows(i)("ProductID") & " AND UnitID = " & dtMainDoc.Rows(i)("UnitID") & " AND " & _
                                "ProductPricePerUnit = " & dtMainDoc.Rows(i)("ProductPricePerUnit")
                    rResult = dtDocRef.Select(strSelect)
                    For j = 0 To rResult.Length - 1
                        dtMainDoc.Rows(i)("UnitSmallAmount") -= rResult(j)("UnitSmallAmount")
                        rResult(j)("AlreadyProcess") = 1
                    Next j
                Next i
                For i = 0 To dtDocRef.Rows.Count - 1
                    If dtDocRef.Rows(i)("AlreadyProcess") = 0 Then
                        dclTemp = dtDocRef.Rows(i)("UnitSmallAmount")
                        strSelect = "ProductID = " & dtDocRef.Rows(i)("ProductID") & " AND UnitID = " & dtDocRef.Rows(i)("UnitID") & " AND UnitSmallAmount > 0 "
                        rResult = dtMainDoc.Select(strSelect)
                        For j = 0 To rResult.Length - 1
                            If rResult(j)("UnitSmallAmount") >= dclTemp Then
                                rResult(j)("UnitSmallAmount") -= dclTemp
                                dclTemp = 0
                            Else
                                dclTemp -= rResult(j)("UnitSmallAmount")
                                rResult(j)("UnitSmallAmount") = 0
                            End If

                            If dclTemp = 0 Then
                                Exit For
                            End If
                        Next j
                        dtDocRef.Rows(i)("AlreadyProcess") = 1
                    End If
                Next i

                Dim strDelDocID As String = ""
                For i = 0 To dtMainDoc.Rows.Count - 1
                    If dtMainDoc.Rows(i)("UnitSmallAmount") <= 0 Then
                        'Delete DocDetail For UnitSmallAmount <=0
                        strDelDocID &= dtMainDoc.Rows(i)("DocDetailID") & ", "
                    Else
                        'Update UnitAmount/ ProductNetPrice/ ProductTax
                        dtMainDoc.Rows(i)("ProductAmount") = dtMainDoc.Rows(i)("UnitSmallAmount") / dtMainDoc.Rows(i)("UnitRatio")
                        dtMainDoc.Rows(i)("ProductNetPrice") = dtMainDoc.Rows(i)("ProductAmount") * dtMainDoc.Rows(i)("ProductPricePerUnit")
                        Select Case dtMainDoc.Rows(i)("ProductTaxType")
                            Case globalVariable.TAXTYPE_INCLUDEVAT
                                dtMainDoc.Rows(i)("ProductTax") = (dtMainDoc.Rows(i)("ProductNetPrice") * globalVariable.DefaultShopVAT) / (100 + globalVariable.DefaultShopVAT)
                                dtMainDoc.Rows(i)("ProductNetPrice") -= dtMainDoc.Rows(i)("ProductTax")
                            Case globalVariable.TAXTYPE_EXCLUDEVAT
                                dtMainDoc.Rows(i)("ProductTax") = (dtMainDoc.Rows(i)("ProductNetPrice") * globalVariable.DefaultShopVAT) / 100
                            Case globalVariable.TAXTYPE_NOVAT
                                dtMainDoc.Rows(i)("ProductTax") = 0
                        End Select
                        DocumentSQL.UpdateMaterialRemainAmountForNewDocumentByDocumentReference(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocID, inventoryID,
                                    dtMainDoc.Rows(i)("DocDetailID"), dtMainDoc.Rows(i)("ProductAmount"), FormatDecimal(dtMainDoc.Rows(i)("ProductTax")),
                                    dtMainDoc.Rows(i)("UnitSmallAmount"), FormatDecimal(dtMainDoc.Rows(i)("ProductNetPrice")), dtMainDoc.Rows(i)("OriginalSmallAmount"))
                    End If
                Next i
                If strDelDocID <> "" Then
                    DocumentSQL.DeleteDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocID, inventoryID, strDelDocID)
                End If

            End If

            Select Case documentTypeID
                Case Is = globalVariable.DOCUMENTTYPE_ROPO
                    DocumentSQL.UpdateDocSummaryIntoDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocID, inventoryID)
                    DocumentSQL.UpdateDocumentStatus(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentRefId, documentRefShopId, globalVariable.DOCUMENTSTATUS_REFERED, strUpdateDate, globalVariable.StaffID)
                Case Is = globalVariable.DOCUMENTTYPE_ROTRANSFER
                    DocumentSQL.UpdateDocumentStatus(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentRefId, documentRefShopId, globalVariable.DOCUMENTSTATUS_APPROVE, strUpdateDate, globalVariable.StaffID, globalVariable.StaffID)
            End Select

            dbTrans.Commit()
            documentData.DocumentID = newDocID
            documentData.DocumentShopID = inventoryID
            documentData.DocumentType = documentTypeID

            Return True
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "DocumentModule", "CreateNewDocumentFromReferedDocument", "99", ex.ToString)
            Return False
        End Try
    End Function

    Friend Function CheckValidDocumentDateForSaveNewDocument(ByVal globalVariable As GlobalVariable, ByVal invendoryId As Integer, ByVal saveDocumentDate As Date, ByRef resultText As String) As Boolean
        Dim lastTransferDay, lastEndDay As Date
        'Check Document Date and Last End Day
        lastEndDay = DocumentModule.GetLastEndDayDocumentDate(globalVariable, invendoryId, lastTransferDay)
        lastTransferDay = DocumentModule.GetLastTransferStockOrCountStockDocumentDate(globalVariable, invendoryId)

        If saveDocumentDate < lastTransferDay Then
            resultText = "ไม่สามารถทำการบันทึกข้อมูลเอกสารนี้ได้ เนื่องจากคลังสินค้าได้มีการยกยอดสต๊อกผ่านเดือนของวันที่ของเอกสารมาแล้ว"
            Return False
        End If
        If saveDocumentDate < lastEndDay Then
            resultText = "ไม่สามารถทำการบันทึกข้อมูลเอกสารนี้ได้ เนื่องจากวันที่ของเอกสารอยู่ก่อนหน้าการทำ End Day ของคลังสินค้านี้ไปแล้ว" & vbLf &
               "(End Day ครั้งล่าสุดเมื่อวันที่ " & Format(lastEndDay, "dd-MMMM-yyyy") & ")"
            Return False
        End If
        resultText = ""
        Return True
    End Function

    Friend Function CalculateSummaryDocDetailPrice(ByVal docDetailList As List(Of DocumentDetail_Data), ByVal docDetailTotalPriceRoundType As RoundType) As DocumentPriceSummary_Data
        Dim docData As DocumentDetail_Data
        Dim DocSummary As New DocumentPriceSummary_Data
        DocSummary.SubTotal = 0
        DocSummary.Discount = 0
        DocSummary.TotalVAT = 0
        DocSummary.NetPrice = 0
        For Each docData In docDetailList
            DocSummary.SubTotal += docData.MaterialTotalPriceBeforeDiscount
            DocSummary.Discount += docData.MaterialDiscount
            DocSummary.TotalVAT += docData.MaterialVAT
            DocSummary.NetPrice += docData.MaterialNetPrice
        Next
        DocSummary.GrandTotal = DocSummary.NetPrice + DocSummary.TotalVAT
        'Rounding Grand Total
        DocSummary.GrandTotal = RoundingPrice(DocSummary.GrandTotal, docDetailTotalPriceRoundType, RoundTo.RoundUp, 0)
        Return DocSummary
    End Function

    Friend Function CalculateSummaryDocDetailPrice(ByVal dtResult As DataTable) As DocumentPriceSummary_Data
        Dim DocSummary As New DocumentPriceSummary_Data
        Dim deliveryVAT, DeliveryNetPrice, deliveryCost As Decimal
        DocSummary.SubTotal = 0
        DocSummary.Discount = 0
        DocSummary.TotalVAT = 0
        DocSummary.NetPrice = 0

        deliveryCost = dtResult.Rows(0)("TransferTotal")
        DeliveryNetPrice = dtResult.Rows(0)("TransferNetPrice")
        deliveryVAT = dtResult.Rows(0)("TransferVAT")

        DocSummary.SubTotal = dtResult.Rows(0)("SubTotal")
        DocSummary.Discount = dtResult.Rows(0)("TotalDiscount")
        DocSummary.Delivery = deliveryCost
        DocSummary.DeliveryVAT = deliveryVAT
        DocSummary.DeliveryNetPrice = DeliveryNetPrice
        DocSummary.TotalVAT = (dtResult.Rows(0)("TotalVAT") + deliveryVAT)
        DocSummary.NetPrice = (dtResult.Rows(0)("NetPrice") + DeliveryNetPrice)
        DocSummary.GrandTotal = (DocSummary.NetPrice + DocSummary.TotalVAT)

        Return DocSummary
    End Function

    Friend Sub CalculateDocDetailAllPrice(ByVal globalVariable As GlobalVariable, ByVal addAmount As Decimal,
                                          ByVal pricePerUnit As Decimal, ByRef discountPercent As Decimal, ByRef discountAmount As Decimal,
                                          ByVal materialVATType As Integer, ByRef docDetailPriceBeforeDiscount As Decimal, ByRef docDetailDiscountPrice As Decimal,
                                          ByRef docDetailTax As Decimal, ByRef docDetailNetPrice As Decimal)

        If addAmount = 0 Then
            discountAmount = 0
        End If
        'Check Valid DiscountPercent
        If discountPercent > 100 Then
            discountPercent = 100
        End If
        docDetailPriceBeforeDiscount = addAmount * pricePerUnit
        'Check Valid Discount Percent
        If discountAmount > docDetailPriceBeforeDiscount Then
            discountAmount = docDetailPriceBeforeDiscount
        End If
        docDetailDiscountPrice = ((docDetailPriceBeforeDiscount * discountPercent) / 100) + discountAmount
        'Total Price --> is price not include VAT
        docDetailNetPrice = docDetailPriceBeforeDiscount - docDetailDiscountPrice
        If pricePerUnit = 0 Then
            docDetailNetPrice = 0
        End If
        'Material Tax
        Select Case materialVATType
            Case globalVariable.TAXTYPE_NOVAT
                docDetailTax = 0
            Case globalVariable.TAXTYPE_EXCLUDEVAT
                docDetailTax = (docDetailNetPrice * globalVariable.DefaultShopVAT) / 100
            Case globalVariable.TAXTYPE_INCLUDEVAT
                docDetailTax = (docDetailNetPrice * globalVariable.DefaultShopVAT) / (100 + globalVariable.DefaultShopVAT)
                docDetailNetPrice -= docDetailTax
        End Select
    End Sub

    Friend Sub CalculateDeliveryCost(ByVal globalVariable As GlobalVariable, ByVal deliveryCost As Decimal, ByVal taxType As Integer, ByRef deliveryTax As Decimal, ByRef deliveryNetPrice As Decimal)

        deliveryNetPrice = deliveryCost
        Select Case taxType
            Case globalVariable.TAXTYPE_NOVAT
                deliveryTax = 0
            Case globalVariable.TAXTYPE_EXCLUDEVAT
                deliveryTax = (deliveryNetPrice * globalVariable.DefaultShopVAT) / 100
                deliveryNetPrice = deliveryCost
            Case globalVariable.TAXTYPE_INCLUDEVAT
                deliveryTax = (deliveryNetPrice * globalVariable.DefaultShopVAT) / (100 + globalVariable.DefaultShopVAT)
                deliveryNetPrice -= deliveryTax
        End Select

    End Sub

    Friend Function CheckWorkingDocument(ByVal globalVariable As GlobalVariable, ByVal documentTypeID As Integer, ByVal documentDate As Date,
                                         ByVal inventoryId As Integer, ByRef documentId As Integer, ByRef documentShopId As Integer,
                                         ByRef resultText As String) As Boolean
        Dim dtResult As New DataTable
        Dim strDocumentDate As String = FormatDate(documentDate)
        dtResult = DocumentSQL.GetDocumentWorkingProcess(globalVariable.DocDBUtil, globalVariable.DocConn, strDocumentDate, documentTypeID, inventoryId, globalVariable.DocLangID)
        If dtResult.Rows.Count > 0 Then
            documentId = dtResult.Rows(0)("documentId")
            documentShopId = dtResult.Rows(0)("shopid")
            Return True
        Else
            documentId = 0
            documentShopId = 0
            Return False
        End If
    End Function

    Friend Function GetDocumentNumberHeader(ByVal dtResult As DataTable, ByVal yearType As Integer, ByRef documentStatus As Integer) As String
        If dtResult.Rows.Count = 0 Then
            documentStatus = 0
            Return ""
        Else
            documentStatus = dtResult.Rows(0)("DocumentStatus")
            If IsDBNull(dtResult.Rows(0)("DocumentYear")) Then
                dtResult.Rows(0)("DocumentYear") = 0
            End If
            If IsDBNull(dtResult.Rows(0)("DocumentMonth")) Then
                dtResult.Rows(0)("DocumentMonth") = 0
            End If
            If IsDBNull(dtResult.Rows(0)("DocumentNumber")) Then
                dtResult.Rows(0)("DocumentNumber") = 0
            End If
            Return GetDocumentHeader(dtResult.Rows(0)("DocumentTypeHeader"), dtResult.Rows(0)("DocumentYear"), _
                dtResult.Rows(0)("DocumentMonth"), dtResult.Rows(0)("DocumentNumber"), yearType)
        End If
    End Function

    Friend Function GetDocumentHeader(ByVal strHeaderType As String, ByVal intYear As Integer, ByVal intMonth As Integer, ByVal strNumber As String, ByVal yearSetting As Integer) As String
        Dim strTemp As String
        Dim strTemp2 As String
        Dim intYearHeader As Integer
        If yearSetting = 1 Then
            intYearHeader = intYear + 543
        Else
            intYearHeader = intYear
        End If
        strTemp2 = intMonth
        If Len(strTemp2) = 1 Then
            strTemp2 = "0" & intMonth
        End If
        strTemp = strNumber
        Do While Len(strTemp) <= 5
            strTemp = "0" & strTemp
        Loop
        strTemp = strHeaderType & strTemp2 & intYearHeader & "/" & strTemp
        Return strTemp
    End Function

    Friend Function GetAndUpdateDocumentNumber(ByVal globalVariable As GlobalVariable, ByVal objTrans As SqlTransaction, ByVal documentShopId As Integer,
                                               ByVal documentTypeId As Integer, ByVal docDate As Date, ByRef docYear As Integer, ByRef docMonth As Integer) As Integer

        Dim newDocNumber, intCurrentYear, intDocMonthInMaxDocumentNumber As Integer
        Dim startIDForNewNumber As Integer
        Dim dtNewID As DataTable
        Dim inVarInfo As System.Globalization.CultureInfo = System.Globalization.CultureInfo.InvariantCulture

        'Get current Christian year
        intCurrentYear = CInt(docDate.ToString("yyyy", inVarInfo))
        intDocMonthInMaxDocumentNumber = docDate.Month
        startIDForNewNumber = 0

        dtNewID = DocumentSQL.GetMaxDocumentNumber(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, documentShopId, documentTypeId, intCurrentYear, intDocMonthInMaxDocumentNumber)
        If dtNewID.Rows.Count = 0 Then
            dtNewID = DocumentSQL.GetMaxDocumentNumberFromDocument(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, documentShopId, documentTypeId, intCurrentYear, intDocMonthInMaxDocumentNumber)
            If dtNewID.Rows(0)("NumRow") = 0 Then
                newDocNumber = 1
            Else
                newDocNumber = dtNewID.Rows(0)("NewID")
            End If
            DocumentSQL.InsertMaxDocumentNumber(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, documentShopId, documentTypeId, intCurrentYear, intDocMonthInMaxDocumentNumber, newDocNumber)
        Else
            newDocNumber = dtNewID.Rows(0)("MaxDocumentNumber") + 1
            DocumentSQL.UpdateMaxDocumentNumber(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, documentShopId, documentTypeId, intCurrentYear, intDocMonthInMaxDocumentNumber, newDocNumber)
        End If
        docYear = intCurrentYear
        docMonth = intDocMonthInMaxDocumentNumber
        Return newDocNumber
    End Function

    Friend Function GetNewDocumentIDFromMaxDocumentID(ByVal globalVariable As GlobalVariable, ByVal objTrans As SqlTransaction, ByVal docShopID As Integer) As Integer
        Dim dtResult As DataTable
        Dim newDocID As Integer
        Dim bolAddNewID As Boolean

        dtResult = DocumentSQL.GetMaxDocumentID(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, docShopID)
        If dtResult.Rows.Count <> 0 Then
            newDocID = dtResult.Rows(0)("MaxDocID") + 1
            bolAddNewID = True
        Else
            dtResult = DocumentSQL.GetMaxDocumentIDFromTableDocument(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, docShopID)
            If dtResult.Rows(0)("CountRow") <> 0 Then
                newDocID = dtResult.Rows(0)("MaxDocID") + 1
            Else
                newDocID = 1
            End If
            bolAddNewID = False
        End If
        If bolAddNewID = False Then
            DocumentSQL.InsertMaxDocumentId(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, docShopID, newDocID)
        Else
            DocumentSQL.UpdateMaxDocumentId(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, docShopID, newDocID)
        End If

        Return newDocID
    End Function

    Friend Function GetLastEndDayDocumentDate(ByVal globalVariable As GlobalVariable, ByVal productLevelID As Integer, ByRef approveStockDate As DateTime) As Date
        Dim strDocType As String
        Dim i As Integer
        Dim dtResult As DataTable

        strDocType = globalVariable.DOCUMENTTYPE_DAILYSTOCK & ", " & globalVariable.DOCUMENTTYPE_WEEKLYSTOCK
        Try
            dtResult = DocumentSQL.GetDocumentTypeStockCountMaterialSetting(globalVariable.DocDBUtil, globalVariable.DocConn, productLevelID, strDocType)
            If dtResult.Rows.Count > 0 Then
                strDocType &= ", "
                For i = 0 To dtResult.Rows.Count - 1
                    strDocType &= dtResult.Rows(i)("DocumentTypeID") & ", "
                Next i
                strDocType = Mid(strDocType, 1, Len(strDocType) - 2)
            End If
        Catch ex As Exception

        End Try

        dtResult = DocumentSQL.GetLastApproveDocumentDate(globalVariable.DocDBUtil, globalVariable.DocConn, productLevelID, strDocType)
        approveStockDate = Date.MinValue
        If dtResult.Rows.Count = 0 Then
            Return Date.MinValue
        ElseIf Not IsDBNull(dtResult.Rows(0)("DocumentDate")) Then
            If Not IsDBNull(dtResult.Rows(0)("ApproveDate")) Then
                approveStockDate = dtResult.Rows(0)("ApproveDate")
            End If
            Return CDate(dtResult.Rows(0)("DocumentDate"))
        Else
            Return Date.MinValue
        End If
    End Function

    Friend Function GetLastTransferStockOrCountStockDocumentDate(ByVal globalVariable As GlobalVariable, ByVal productLevelID As Integer) As Date
        Dim dtResult As DataTable
        dtResult = DocumentSQL.GetLastTransferStockOrCountStockDocumentDate(globalVariable.DocDBUtil, globalVariable.DocConn, productLevelID)
        If dtResult.Rows.Count = 0 Then
            Return Date.MinValue
        ElseIf Not IsDBNull(dtResult.Rows(0)("DocumentDate")) Then
            Return CDate(dtResult.Rows(0)("DocumentDate"))
        Else
            Return Date.MinValue
        End If
    End Function

    Friend Function ConvertAmountFromSmallToLargeAmount(ByVal smallAmount As Decimal, ByVal dtAllRatio As DataTable, ByVal unitSmallID As Integer, ByVal toUnitLargeID As Integer) As Decimal
        Dim rResult() As DataRow
        rResult = dtAllRatio.Select("UnitSmallID = " & unitSmallID & " AND UnitLargeID = " & toUnitLargeID)
        If rResult.Length = 0 Then
            Return smallAmount
        Else
            If rResult(0)("UnitSmallRatio") = 0 Then
                rResult(0)("UnitSmallRatio") = 1
            End If
            Return Format((smallAmount * rResult(0)("UnitLargeRatio")) / rResult(0)("UnitSmallRatio"), "0.0000")
        End If
    End Function

    Friend Function InsertResultDataIntoList(ByVal globalVariable As GlobalVariable, ByVal searchDocType As Integer, ByVal dtSearchResult As DataTable, ByVal dtDocStatus As DataTable) As List(Of SearchDocumentResult_Data)


        Dim dueDate As Date = Date.MinValue
        Dim strDocNumber As String = ""
        Dim strDocRefNumber As String = ""
        Dim docReceiveBy As Integer = 0
        Dim docToInventoryID As Integer = 0
        Dim docFromInventoryID As Integer = 0
        Dim refDocumentType As Integer = 0
        Dim refDocumentTypeName As String = ""
        Dim refDocumentTypeHeader As String = ""
        Dim refDocumentStatus As Integer = 0
        Dim refDocumentYear As Integer = 0
        Dim refDocumentMonth As Integer = 0
        Dim refDocumentNumber As Integer = 0
        Dim documentIDRef As Integer = 0
        Dim documentIDRefShopID As Integer = 0
        Dim dtInventoryData As New DataTable
        Dim toInventoryName As String = ""
        Dim fromInventoryName As String = ""
        Dim resultList As New List(Of SearchDocumentResult_Data)
        resultList = New List(Of SearchDocumentResult_Data)
        Dim dtUser As New DataTable
        Dim dtBusinessPlace As New DataTable
        Dim businessPlaceName As String = ""

        dtUser = DocumentSQL.GetUser(globalVariable.DocDBUtil, globalVariable.DocConn)
        dtInventoryData = InventorySQL.GetInventory(globalVariable.DocDBUtil, globalVariable.DocConn, 1)
        dtBusinessPlace = DocumentSQL.ListBusinessPlace(globalVariable.DocDBUtil, globalVariable.DocConn)

        For i = 0 To dtSearchResult.Rows.Count - 1
            If IsDBNull(dtSearchResult.Rows(i)("DocumentTypeName")) Then
                dtSearchResult.Rows(i)("DocumentTypeName") = ""
            End If
            'Insert DocumentNumber
            If Not IsDBNull(dtSearchResult.Rows(i)("DocumentTypeHeader")) Then
                strDocNumber = dtSearchResult.Rows(i)("DocumentTypeHeader")
            Else
                strDocNumber = ""
            End If
            strDocNumber = GetDocumentHeader(strDocNumber, dtSearchResult.Rows(i)("DocumentYear"), dtSearchResult.Rows(i)("DocumentMonth"), dtSearchResult.Rows(i)("DocumentNumber"), globalVariable.DocYearSettingType)

            If IsDBNull(dtSearchResult.Rows(i)("VendorCode")) Then
                dtSearchResult.Rows(i)("VendorCode") = ""
            End If
            If IsDBNull(dtSearchResult.Rows(i)("VendorName")) Then
                dtSearchResult.Rows(i)("VendorName") = ""
            End If
            If IsDBNull(dtSearchResult.Rows(i)("Remark")) Then
                dtSearchResult.Rows(i)("Remark") = ""
            End If
            If IsDBNull(dtSearchResult.Rows(i)("InvoiceRef")) Then
                dtSearchResult.Rows(i)("InvoiceRef") = ""
            End If
            'Reference Document
            If dtSearchResult.Columns.Contains("RefDocumentType") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("documentIDRef")) Then
                    documentIDRef = dtSearchResult.Rows(i)("documentIDRef")
                End If
            End If
            If dtSearchResult.Columns.Contains("documentIDRefShopID") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("documentIDRefShopID")) Then
                    documentIDRefShopID = dtSearchResult.Rows(i)("documentIDRefShopID")
                End If
            End If
            If dtSearchResult.Columns.Contains("RefDocumentType") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("RefDocumentType")) Then
                    refDocumentType = dtSearchResult.Rows(i)("RefDocumentType")
                End If
            End If
            If dtSearchResult.Columns.Contains("RefDocumentTypeName") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("RefDocumentTypeName")) Then
                    refDocumentTypeName = dtSearchResult.Rows(i)("RefDocumentTypeName")
                End If
            End If
            If dtSearchResult.Columns.Contains("RefDocumentTypeHeader") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("RefDocumentTypeHeader")) Then
                    refDocumentTypeHeader = dtSearchResult.Rows(i)("RefDocumentTypeHeader")
                End If
            End If
            If dtSearchResult.Columns.Contains("RefDocumentStatus") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("RefDocumentStatus")) Then
                    refDocumentStatus = dtSearchResult.Rows(i)("RefDocumentStatus")
                End If
            End If
            If dtSearchResult.Columns.Contains("RefDocumentYear") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("RefDocumentYear")) Then
                    refDocumentYear = dtSearchResult.Rows(i)("RefDocumentYear")
                End If
            End If
            If dtSearchResult.Columns.Contains("RefDocumentMonth") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("RefDocumentMonth")) Then
                    refDocumentMonth = dtSearchResult.Rows(i)("RefDocumentMonth") = 0
                End If
            End If
            If dtSearchResult.Columns.Contains("RefDocumentNumber") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("RefDocumentNumber")) Then
                    refDocumentNumber = dtSearchResult.Rows(i)("RefDocumentNumber")

                    strDocRefNumber = GetDocumentHeader(dtSearchResult.Rows(i)("RefDocumentTypeHeader"), dtSearchResult.Rows(i)("RefDocumentYear"), dtSearchResult.Rows(i)("RefDocumentMonth"), dtSearchResult.Rows(i)("RefDocumentNumber"), globalVariable.DocYearSettingType)
                Else
                    strDocRefNumber = ""
                End If
            End If
            If dtSearchResult.Columns.Contains("DueDate") = True Then
                If Not IsDBNull(dtSearchResult.Rows(i)("DueDate")) Then
                    dueDate = dtSearchResult.Rows(i)("DueDate")
                Else
                    dueDate = Date.MinValue
                End If
            End If

            Dim statusName As String = ""
            Dim expression As String = ""
            Dim foundRows() As DataRow

            expression = "StatusId=" & dtSearchResult.Rows(i)("DocumentStatus")
            foundRows = dtDocStatus.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                statusName = foundRows(0)("StatusName")
            End If

            expression = "ShopID=" & dtSearchResult.Rows(i)("ToInvID")
            foundRows = dtInventoryData.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                If Not IsDBNull(foundRows(0)("ShopID")) Then
                    toInventoryName = foundRows(0)("ShopID")
                Else
                    toInventoryName = ""
                End If
            End If

            expression = "ShopID=" & dtSearchResult.Rows(i)("ToInvID")
            foundRows = dtInventoryData.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                If Not IsDBNull(foundRows(0)("ShopName")) Then
                    toInventoryName = foundRows(0)("ShopName")
                Else
                    toInventoryName = ""
                End If
            End If
            If searchDocType = 3 Then
                expression = "ShopID=" & dtSearchResult.Rows(i)("ProductLevelID")
                foundRows = dtInventoryData.Select(expression)
                If foundRows.GetUpperBound(0) >= 0 Then
                    If Not IsDBNull(foundRows(0)("ShopName")) Then
                        fromInventoryName = foundRows(0)("ShopName")
                    Else
                        fromInventoryName = ""
                    End If
                End If
            Else
                expression = "ShopID=" & dtSearchResult.Rows(i)("FromInvID")
                foundRows = dtInventoryData.Select(expression)
                If foundRows.GetUpperBound(0) >= 0 Then
                    If Not IsDBNull(foundRows(0)("ShopName")) Then
                        fromInventoryName = foundRows(0)("ShopName")
                    Else
                        fromInventoryName = ""
                    End If
                End If
            End If

            Dim InsertStaffName As String = ""
            Dim UpdateStaffName As String = ""
            Dim ApproveStaffName As String = ""
            Dim CancelStaffName As String = ""

            expression = "StaffId=" & dtSearchResult.Rows(i)("InputBy")
            foundRows = dtUser.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                If Not IsDBNull(foundRows(0)("StaffName")) Then
                    InsertStaffName = foundRows(0)("StaffName")
                End If
            End If

            expression = "StaffId=" & dtSearchResult.Rows(i)("UpdateBy")
            foundRows = dtUser.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                If Not IsDBNull(foundRows(0)("StaffName")) Then
                    UpdateStaffName = foundRows(0)("StaffName")
                End If
            End If

            expression = "StaffId=" & dtSearchResult.Rows(i)("ApproveBy")
            foundRows = dtUser.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                If Not IsDBNull(foundRows(0)("StaffName")) Then
                    ApproveStaffName = foundRows(0)("StaffName")
                End If
            End If

            expression = "StaffId=" & dtSearchResult.Rows(i)("VoidBy")
            foundRows = dtUser.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                If Not IsDBNull(foundRows(0)("StaffName")) Then
                    CancelStaffName = foundRows(0)("StaffName")
                End If
            End If
            expression = "BUS_ID=" & dtSearchResult.Rows(i)("BusinessPlace")
            foundRows = dtBusinessPlace.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                businessPlaceName = foundRows(0)("BUS_NAME")
            Else
                businessPlaceName = ""
            End If
            Dim taxInvoiceDate As Date
            If Not IsDBNull(dtSearchResult.Rows(i)("TaxInvoiceDate")) Then
                taxInvoiceDate = dtSearchResult.Rows(i)("TaxInvoiceDate")
            Else
                taxInvoiceDate = Date.MinValue
            End If
            If Not IsDBNull(dtSearchResult.Rows(i)("TaxInvoiceNo")) Then
                dtSearchResult.Rows(i)("TaxInvoiceNo") = dtSearchResult.Rows(i)("TaxInvoiceNo")
            Else
                dtSearchResult.Rows(i)("TaxInvoiceNo") = ""
            End If
            resultList.Add(SearchDocumentResult_Data.NewSearchDocumentResult(dtSearchResult.Rows(i)("DocumentID"), dtSearchResult.Rows(i)("ShopID"),
                    dtSearchResult.Rows(i)("DocumentStatus"), statusName, docReceiveBy, searchDocType, dtSearchResult.Rows(i)("DocumentTypeName"),
                    documentIDRef, documentIDRefShopID, refDocumentStatus, refDocumentType, refDocumentTypeName, dtSearchResult.Rows(i)("DocumentDate"), dueDate,
                    strDocNumber, strDocRefNumber, dtSearchResult.Rows(i)("ShopId"), dtSearchResult.Rows(i)("ShopId"), dtSearchResult.Rows(i)("ToInvID"),
                    toInventoryName, dtSearchResult.Rows(i)("FromInvID"), fromInventoryName, dtSearchResult.Rows(i)("Remark"), dtSearchResult.Rows(i)("InvoiceRef"),
                    dtSearchResult.Rows(i)("VendorID"), dtSearchResult.Rows(i)("VendorGroupID"), dtSearchResult.Rows(i)("VendorShopID"), dtSearchResult.Rows(i)("VendorCode"),
                    dtSearchResult.Rows(i)("VendorName"), dtSearchResult.Rows(i)("SubTotal"), dtSearchResult.Rows(i)("TotalDiscount"), dtSearchResult.Rows(i)("TotalVAT"),
                    dtSearchResult.Rows(i)("NetPrice"), dtSearchResult.Rows(i)("GrandTotal"), InsertStaffName, UpdateStaffName, ApproveStaffName, CancelStaffName, businessPlaceName,
                    dtSearchResult.Rows(i)("TaxInvoiceNo"), taxInvoiceDate))
        Next i
        Return resultList
    End Function

    Friend Function LoadDocument(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer, ByRef docData As Document_Data,
                                 ByRef resultText As String) As Boolean

        Dim dtResult As New DataTable
        Dim dtDocRef As New DataTable
        Dim dtDocStatus As New DataTable
        Dim dtUser As New DataTable
        Dim dtVendor As New DataTable
        Dim tempDate As DateTime
        Dim expression As String = ""
        Dim foundRows() As DataRow

        docData = New Document_Data
        dtResult = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId, globalVariable.DocLangID)
        dtDocStatus = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
        dtUser = DocumentSQL.GetUser(globalVariable.DocDBUtil, globalVariable.DocConn)

        'No Document
        If dtResult.Rows.Count = 0 Then
            docData.DocumentID = -1
            docData.DocumentShopID = -1
            resultText = globalVariable.MESSAGE_DATANOTFOUND
            Return False
        End If
        dtVendor = VendorSQL.GetVendorDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dtResult.Rows(0)("VendorId"), globalVariable.DocLangID)

        docData.SoftwareVersion = globalVariable.SoftwareVersion
        'Set DocumentType/ DocumentStatus/ DocumentHeader 
        docData.DocumentType = dtResult.Rows(0)("DocumentTypeID")
        docData.DocumentTypeName = dtResult.Rows(0)("DocumentTypeName")
        docData.DocumentStatus = dtResult.Rows(0)("DocumentStatus")

        expression = "StatusId=" & dtResult.Rows(0)("DocumentStatus")
        foundRows = dtDocStatus.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StatusName")) Then
                docData.DocumentStatusName = foundRows(0)("StatusName")
            End If
        Else
            docData.DocumentStatusName = ""
        End If
        If dtResult.Rows(0)("DocumentStatus") = 0 Then
            docData.DocumentNumber = ""
        Else
            docData.DocumentNumber = GetDocumentHeader(dtResult.Rows(0)("DocumentTypeHeader"), dtResult.Rows(0)("DocumentYear"), dtResult.Rows(0)("DocumentMonth"), dtResult.Rows(0)("DocumentNumber"), globalVariable.DocYearSettingType)
        End If
        If dtResult.Rows(0)("DocumentIDRef") <> 0 Then
            dtDocRef = DocumentSQL.GetDocumentNumber(globalVariable.DocDBUtil, globalVariable.DocConn, dtResult.Rows(0)("DocumentIDRef"), dtResult.Rows(0)("DocumentIDRefShopID"), globalVariable.DocLangID)
            docData.DocumentRefNumber = GetDocumentHeader(dtDocRef.Rows(0)("DocumentTypeHeader"), dtDocRef.Rows(0)("DocumentYear"), dtDocRef.Rows(0)("DocumentMonth"), dtDocRef.Rows(0)("DocumentNumber"), globalVariable.DocYearSettingType)
            docData.DocumentRefID = dtResult.Rows(0)("DocumentIDRef")
            docData.DocumentRefShopID = dtResult.Rows(0)("DocumentIDRefShopID")
        Else
            docData.DocumentRefNumber = ""
            docData.DocumentRefID = 0
            docData.DocumentRefShopID = 0
            docData.DocumentRefStatus = 0
        End If
        If Not IsDBNull(dtResult.Rows(0)("DocumentDate")) Then
            tempDate = dtResult.Rows(0)("DocumentDate")
            docData.DocumentDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        If Not IsDBNull(dtResult.Rows(0)("InvoiceRef")) Then
            docData.InvoiceRef = dtResult.Rows(0)("InvoiceRef")
        Else
            docData.InvoiceRef = ""
        End If
        If Not IsDBNull(dtResult.Rows(0)("Remark")) Then
            docData.DocumentNote = dtResult.Rows(0)("Remark")
        Else
            docData.DocumentNote = ""
        End If
        docData.DocumentInventoryID = dtResult.Rows(0)("ProductLevelID")
        docData.DocumentToInventoryID = dtResult.Rows(0)("ToInvID")
        docData.DocumentFromInventoryID = dtResult.Rows(0)("FromInvID")
        If dtResult.Rows(0)("TransferTotal") > 0 Then
            docData.DeliveryCost = dtResult.Rows(0)("TransferTotal")
        End If
        docData.DeliveryCostVAT = dtResult.Rows(0)("TransferVAT")
        docData.DeliveryCostNetPrice = dtResult.Rows(0)("TransferNetPrice")
        'Set Vendor
        docData.VendorID = dtResult.Rows(0)("VendorID")
        docData.VendorGroupID = dtResult.Rows(0)("VendorGroupID")
        docData.VendorGroupShopID = dtResult.Rows(0)("VendorShopID")

        'Can Edit DocDetail Amount
        If dtResult.Rows(0)("LockEditDetail") = 1 Then
            docData.LockEditDocDetail = True
        Else
            docData.LockEditDocDetail = False
        End If

        'Term of Payment/ Credit Day/ DueDate
        If IsDBNull(dtResult.Rows(0)("TermOfPayment")) Then
            docData.TermOfPayment = 0
        Else
            If dtResult.Rows(0)("TermOfPayment") > 0 Then
                docData.TermOfPayment = dtResult.Rows(0)("TermOfPayment")
            End If
        End If
        If dtResult.Rows(0)("CreditDay") > 0 Then
            docData.CreditDay = dtResult.Rows(0)("CreditDay")
        End If

        'DueDate
        If Not IsDBNull(dtResult.Rows(0)("DueDate")) Then
            tempDate = dtResult.Rows(0)("DueDate")
            If tempDate <> Date.MinValue Then
                docData.DueDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
                docData.DeliveryTime = tempDate.ToString("HH:mm", globalVariable.InvariantCulture)
            End If
        End If

        'PO Invoice Date/ Price
        If Not IsDBNull(dtResult.Rows(0)("InvoicePODate")) Then
            tempDate = dtResult.Rows(0)("InvoicePODate")
            docData.InvoicePODate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        docData.InvoicePOTotalPriceBeforeVAT = dtResult.Rows(0)("InvoicePOTotalPriceBeforeVAT")
        docData.InvoicePOTotalPriceIncludeVAT = dtResult.Rows(0)("InvoicePOTotalPriceIncludeVAT")

        'Set Insert/ Update/ Approve/ CancelDate
        If Not IsDBNull(dtResult.Rows(0)("InsertDate")) Then
            tempDate = dtResult.Rows(0)("InsertDate")
            docData.InsertDate = tempDate.ToString("yyyy-MM-dd HH:mm:ss", globalVariable.InvariantCulture)
        End If
        docData.InsertStaffID = dtResult.Rows(0)("InputBy")
        expression = "StaffId=" & dtResult.Rows(0)("InputBy")
        foundRows = dtUser.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StaffName")) Then
                docData.InsertStaffName = foundRows(0)("StaffName")
            End If
        End If
        If Not IsDBNull(dtResult.Rows(0)("UpdateDate")) Then
            tempDate = dtResult.Rows(0)("UpdateDate")
            docData.UpdateDate = tempDate.ToString("yyyy-MM-dd HH:mm:ss", globalVariable.InvariantCulture)
        End If

        docData.UpdateStaffID = dtResult.Rows(0)("UpdateBy")
        expression = "StaffId=" & dtResult.Rows(0)("UpdateBy")
        foundRows = dtUser.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StaffName")) Then
                docData.UpdateStaffName = foundRows(0)("StaffName")
            End If
        End If
        If Not IsDBNull(dtResult.Rows(0)("ApproveDate")) Then
            tempDate = dtResult.Rows(0)("ApproveDate")
            docData.ApproveDate = tempDate.ToString("yyyy-MM-dd HH:mm:ss", globalVariable.InvariantCulture)
        End If

        docData.ApproveDocStaffID = dtResult.Rows(0)("ApproveBy")
        expression = "StaffId=" & dtResult.Rows(0)("ApproveBy")
        foundRows = dtUser.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StaffName")) Then
                docData.ApproveDocStaffName = foundRows(0)("StaffName")
            End If
        End If
        If Not IsDBNull(dtResult.Rows(0)("CancelDate")) Then
            tempDate = dtResult.Rows(0)("CancelDate")
            docData.CancelDate = tempDate.ToString("yyyy-MM-dd HH:mm:ss", globalVariable.InvariantCulture)
        End If
        docData.CancelStaffID = dtResult.Rows(0)("VoidBy")
        expression = "StaffId=" & dtResult.Rows(0)("VoidBy")
        foundRows = dtUser.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StaffName")) Then
                docData.CancelStaffName = foundRows(0)("StaffName")
            End If
        End If
        If dtResult.Rows(0)("MovementInStock") = -1 Then
            docData.IsAddReduceDoc = 2
        Else
            docData.IsAddReduceDoc = 1
        End If
        If Not IsDBNull(dtResult.Rows(0)("StockAtDateTime")) Then
            tempDate = dtResult.Rows(0)("StockAtDateTime")
            docData.StockAtDateTime = tempDate.ToString("yyyy-MM-dd HH:mm:ss", globalVariable.InvariantCulture)
        End If
        If dtVendor.Rows.Count > 0 Then
            docData.DefaultTaxType = dtVendor.Rows(0)("DefaultTaxType")
        End If
        docData.LastTransferStock = CountStockModule.GetLastTransferStock(globalVariable, documentShopId)
        LoadDocumentDetail(globalVariable, documentId, documentShopId, docData, resultText)
        docData.DocSummary = CalculateSummaryDocDetailPrice(dtResult)
        'Set Current DocumentID
        docData.DocumentID = documentId
        docData.DocumentShopID = documentShopId

        resultText = ""
        Return True
    End Function

    Friend Function LoadDocument(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer, ByRef docData As DocumentPTT_Data,
                               ByRef resultText As String) As Boolean

        Dim dtResult As New DataTable
        Dim dtDocRef As New DataTable
        Dim dtDocStatus As New DataTable
        Dim dtUser As New DataTable
        Dim dtVendor As New DataTable
        Dim tempDate As DateTime
        Dim expression As String = ""
        Dim foundRows() As DataRow

        docData = New DocumentPTT_Data
        dtResult = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId, globalVariable.DocLangID)
        dtDocStatus = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
        dtUser = DocumentSQL.GetUser(globalVariable.DocDBUtil, globalVariable.DocConn)

        'No Document
        If dtResult.Rows.Count = 0 Then
            docData.DocumentID = -1
            docData.DocumentShopID = -1
            resultText = globalVariable.MESSAGE_DATANOTFOUND
            Return False
        End If
        dtVendor = VendorSQL.GetVendorDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dtResult.Rows(0)("VendorId"), globalVariable.DocLangID)
        docData.SoftwareVersion = globalVariable.SoftwareVersion
        'Set DocumentType/ DocumentStatus/ DocumentHeader 
        docData.DocumentType = dtResult.Rows(0)("DocumentTypeID")
        docData.DocumentTypeName = dtResult.Rows(0)("DocumentTypeName")
        docData.DocumentStatus = dtResult.Rows(0)("DocumentStatus")

        expression = "StatusId=" & dtResult.Rows(0)("DocumentStatus")
        foundRows = dtDocStatus.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StatusName")) Then
                docData.DocumentStatusName = foundRows(0)("StatusName")
            End If
        Else
            docData.DocumentStatusName = ""
        End If
        If dtResult.Rows(0)("DocumentStatus") = 0 Then
            docData.DocumentNumber = ""
        Else
            docData.DocumentNumber = GetDocumentHeader(dtResult.Rows(0)("DocumentTypeHeader"), dtResult.Rows(0)("DocumentYear"), dtResult.Rows(0)("DocumentMonth"), dtResult.Rows(0)("DocumentNumber"), globalVariable.DocYearSettingType)
        End If
        If dtResult.Rows(0)("DocumentIDRef") <> 0 Then
            dtDocRef = DocumentSQL.GetDocumentNumber(globalVariable.DocDBUtil, globalVariable.DocConn, dtResult.Rows(0)("DocumentIDRef"), dtResult.Rows(0)("DocumentIDRefShopID"), globalVariable.DocLangID)
            docData.DocumentRefNumber = GetDocumentHeader(dtDocRef.Rows(0)("DocumentTypeHeader"), dtDocRef.Rows(0)("DocumentYear"), dtDocRef.Rows(0)("DocumentMonth"), dtDocRef.Rows(0)("DocumentNumber"), globalVariable.DocYearSettingType)
            docData.DocumentRefID = dtResult.Rows(0)("DocumentIDRef")
            docData.DocumentRefShopID = dtResult.Rows(0)("DocumentIDRefShopID")
        Else
            docData.DocumentRefNumber = ""
            docData.DocumentRefID = 0
            docData.DocumentRefShopID = 0
            docData.DocumentRefStatus = 0
        End If
        If Not IsDBNull(dtResult.Rows(0)("DocumentDate")) Then
            tempDate = dtResult.Rows(0)("DocumentDate")
            docData.DocumentDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        If Not IsDBNull(dtResult.Rows(0)("InvoiceRef")) Then
            docData.InvoiceRef = dtResult.Rows(0)("InvoiceRef")
        Else
            docData.InvoiceRef = ""
        End If
        If Not IsDBNull(dtResult.Rows(0)("Remark")) Then
            docData.DocumentNote = dtResult.Rows(0)("Remark")
        Else
            docData.DocumentNote = ""
        End If
        docData.DocumentInventoryID = dtResult.Rows(0)("ProductLevelID")
        docData.DocumentToInventoryID = dtResult.Rows(0)("ToInvID")
        docData.DocumentFromInventoryID = dtResult.Rows(0)("FromInvID")

        'Set Vendor
        docData.VendorID = dtResult.Rows(0)("VendorID")
        docData.VendorGroupID = dtResult.Rows(0)("VendorGroupID")
        docData.VendorGroupShopID = dtResult.Rows(0)("VendorShopID")

        'Can Edit DocDetail Amount
        If dtResult.Rows(0)("LockEditDetail") = 1 Then
            docData.LockEditDocDetail = True
        Else
            docData.LockEditDocDetail = False
        End If

        'Term of Payment/ Credit Day/ DueDate
        If IsDBNull(dtResult.Rows(0)("TermOfPayment")) Then
            docData.TermOfPayment = 0
        Else
            If dtResult.Rows(0)("TermOfPayment") > 0 Then
                docData.TermOfPayment = dtResult.Rows(0)("TermOfPayment")
            End If
        End If
        If dtResult.Rows(0)("CreditDay") > 0 Then
            docData.CreditDay = dtResult.Rows(0)("CreditDay")
        End If
       
        'DueDate
        If Not IsDBNull(dtResult.Rows(0)("DueDate")) Then
            tempDate = dtResult.Rows(0)("DueDate")
            If tempDate <> Date.MinValue Then
                docData.DueDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
            End If
        End If

        'PO Invoice Date/ Price
        If Not IsDBNull(dtResult.Rows(0)("InvoicePODate")) Then
            tempDate = dtResult.Rows(0)("InvoicePODate")
            docData.InvoicePODate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        docData.InvoicePOTotalPriceBeforeVAT = dtResult.Rows(0)("InvoicePOTotalPriceBeforeVAT")
        docData.InvoicePOTotalPriceIncludeVAT = dtResult.Rows(0)("InvoicePOTotalPriceIncludeVAT")

        'Set Insert/ Update/ Approve/ CancelDate
        If Not IsDBNull(dtResult.Rows(0)("InsertDate")) Then
            tempDate = dtResult.Rows(0)("InsertDate")
            docData.InsertDate = tempDate.ToString("yyyy-MM-dd HH:mm:ss", globalVariable.InvariantCulture)
        End If
        docData.InsertStaffID = dtResult.Rows(0)("InputBy")
        expression = "StaffId=" & dtResult.Rows(0)("InputBy")
        foundRows = dtUser.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StaffName")) Then
                docData.InsertStaffName = foundRows(0)("StaffName")
            End If
        End If
        If Not IsDBNull(dtResult.Rows(0)("UpdateDate")) Then
            tempDate = dtResult.Rows(0)("UpdateDate")
            docData.UpdateDate = tempDate.ToString("yyyy-MM-dd HH:mm:ss", globalVariable.InvariantCulture)
        End If

        docData.UpdateStaffID = dtResult.Rows(0)("UpdateBy")
        expression = "StaffId=" & dtResult.Rows(0)("UpdateBy")
        foundRows = dtUser.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StaffName")) Then
                docData.UpdateStaffName = foundRows(0)("StaffName")
            End If
        End If
        If Not IsDBNull(dtResult.Rows(0)("ApproveDate")) Then
            tempDate = dtResult.Rows(0)("ApproveDate")
            docData.ApproveDate = tempDate.ToString("yyyy-MM-dd HH:mm:ss", globalVariable.InvariantCulture)
        End If

        docData.ApproveDocStaffID = dtResult.Rows(0)("ApproveBy")
        expression = "StaffId=" & dtResult.Rows(0)("ApproveBy")
        foundRows = dtUser.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StaffName")) Then
                docData.ApproveDocStaffName = foundRows(0)("StaffName")
            End If
        End If
        If Not IsDBNull(dtResult.Rows(0)("CancelDate")) Then
            tempDate = dtResult.Rows(0)("CancelDate")
            docData.CancelDate = tempDate.ToString("yyyy-MM-dd HH:mm:ss", globalVariable.InvariantCulture)
        End If
        docData.CancelStaffID = dtResult.Rows(0)("VoidBy")
        expression = "StaffId=" & dtResult.Rows(0)("VoidBy")
        foundRows = dtUser.Select(expression)
        If foundRows.GetUpperBound(0) >= 0 Then
            If Not IsDBNull(foundRows(0)("StaffName")) Then
                docData.CancelStaffName = foundRows(0)("StaffName")
            End If
        End If
        If Not IsDBNull(dtResult.Rows(0)("CustomerDocNo")) Then
            docData.CustomerDocNo = dtResult.Rows(0)("CustomerDocNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("CustomerCode")) Then
            docData.CustomerCode = dtResult.Rows(0)("CustomerCode")
        End If
        If Not IsDBNull(dtResult.Rows(0)("CustomerName")) Then
            docData.CustomerName = dtResult.Rows(0)("CustomerName")
        End If
        If Not IsDBNull(dtResult.Rows(0)("CustomerAddress")) Then
            docData.CustomerAddress = dtResult.Rows(0)("CustomerAddress")
        End If
        If Not IsDBNull(dtResult.Rows(0)("ShipToOrDestinationPort")) Then
            docData.CustomerShipTo = dtResult.Rows(0)("ShipToOrDestinationPort")
        End If
        If Not IsDBNull(dtResult.Rows(0)("BillTo")) Then
            docData.CustomerBillTo = dtResult.Rows(0)("BillTo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("TaxInvoiceNo")) Then
            docData.TaxInvoiceNo = dtResult.Rows(0)("TaxInvoiceNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("TaxInvoiceDate")) Then
            tempDate = dtResult.Rows(0)("TaxInvoiceDate")
            docData.TaxInvoiceDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        If Not IsDBNull(dtResult.Rows(0)("invoiceNo")) Then
            docData.InvoiceNo = dtResult.Rows(0)("invoiceNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("invoiceDate")) Then
            tempDate = dtResult.Rows(0)("invoiceDate")
            docData.InvoiceDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        If Not IsDBNull(dtResult.Rows(0)("saleOrderNo")) Then
            docData.SaleOrderNo = dtResult.Rows(0)("saleOrderNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("SaleOrderDate")) Then
            tempDate = dtResult.Rows(0)("SaleOrderDate")
            docData.InvoiceDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        If Not IsDBNull(dtResult.Rows(0)("InvoicePONO")) Then
            docData.PurchaseOrderNo = dtResult.Rows(0)("InvoicePONO")
        End If
        If Not IsDBNull(dtResult.Rows(0)("InvoicePODate")) Then
            tempDate = dtResult.Rows(0)("InvoicePODate")
            docData.PurchaseOrderDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        If Not IsDBNull(dtResult.Rows(0)("DeliveryNo")) Then
            docData.DeliveryNo = dtResult.Rows(0)("DeliveryNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("DeliveryDate")) Then
            tempDate = dtResult.Rows(0)("DeliveryDate")
            docData.DeliveryDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        If Not IsDBNull(dtResult.Rows(0)("BusinessPlace")) Then
            docData.BusinessID = dtResult.Rows(0)("BusinessPlace")
        End If
        If Not IsDBNull(dtResult.Rows(0)("Plant")) Then
            docData.PlantID = dtResult.Rows(0)("Plant")
        End If
        If Not IsDBNull(dtResult.Rows(0)("PaymentDate")) Then
            tempDate = dtResult.Rows(0)("PaymentDate")
            docData.PaymentDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        If Not IsDBNull(dtResult.Rows(0)("Incoterm")) Then
            docData.Incoterm = dtResult.Rows(0)("Incoterm")
        End If
        If Not IsDBNull(dtResult.Rows(0)("ContractNo")) Then
            docData.ContractNo = dtResult.Rows(0)("ContractNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("ShipmentDate")) Then
            tempDate = dtResult.Rows(0)("ShipmentDate")
            docData.ShipmentDate = tempDate.ToString("yyyy-MM-dd", globalVariable.InvariantCulture)
        End If
        If Not IsDBNull(dtResult.Rows(0)("ShippingCondition")) Then
            docData.ShippingCondition = dtResult.Rows(0)("ShippingCondition")
        End If
        If Not IsDBNull(dtResult.Rows(0)("CarrierBy")) Then
            docData.CarrierBy = dtResult.Rows(0)("CarrierBy")
        End If
        If Not IsDBNull(dtResult.Rows(0)("CarrierNo")) Then
            docData.CarrierNo = dtResult.Rows(0)("CarrierNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("DriverName")) Then
            docData.DriverName = dtResult.Rows(0)("DriverName")
        End If
        If Not IsDBNull(dtResult.Rows(0)("SealNo")) Then
            docData.SealNo = dtResult.Rows(0)("SealNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("TripNo")) Then
            docData.TripNo = dtResult.Rows(0)("TripNo")
        End If
        docData.ShippingCostTaxType = dtResult.Rows(0)("TransferTaxClass")
        docData.DeliveryCostTaxType = dtResult.Rows(0)("TransferTaxClass")
        If dtResult.Rows(0)("TransferTotal") > 0 Then
            docData.ShippingCost = dtResult.Rows(0)("TransferTotal")
        End If
        If dtResult.Rows(0)("TransferTotal") > 0 Then
            docData.DeliveryCost = dtResult.Rows(0)("TransferTotal")
        End If
        docData.DeliveryCostVAT = dtResult.Rows(0)("TransferVAT")
        docData.DeliveryCostNetPrice = dtResult.Rows(0)("TransferNetPrice")
        If Not IsDBNull(dtResult.Rows(0)("ShipmentNo")) Then
            docData.ShipmentNo = dtResult.Rows(0)("ShipmentNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("ShipmentNo")) Then
            docData.ShipmentNo = dtResult.Rows(0)("ShipmentNo")
        End If
        If Not IsDBNull(dtResult.Rows(0)("GSNO")) Then
            docData.GS_No = dtResult.Rows(0)("GSNO")
        End If

        docData.ShiftID = dtResult.Rows(0)("ShiftID")
        docData.ShiftNo = dtResult.Rows(0)("ShiftNo")
        docData.ShiftDay = dtResult.Rows(0)("ShiftDay")

        If Not IsDBNull(dtResult.Rows(0)("CreateBy")) Then
            docData.CreateBy = dtResult.Rows(0)("CreateBy")
        End If
        If dtVendor.Rows.Count > 0 Then
            docData.DefaultTaxType = dtVendor.Rows(0)("DefaultTaxType")
        End If
        LoadDocumentDetail(globalVariable, documentId, documentShopId, docData, resultText)
        docData.DocSummary = CalculateSummaryDocDetailPrice(dtResult)
        Dim tankDetail As New List(Of DocumentDetailTank_Data)
        LoadDocumentDetailTank(globalVariable, documentId, documentShopId, tankDetail, resultText)
        docData.DocDetailTankList = tankDetail
        'Set Current DocumentID
        docData.DocumentID = documentId
        docData.DocumentShopID = documentShopId

        resultText = ""
        Return True
    End Function

    Friend Function LoadDocumentDetail(ByVal globalVariable As GlobalVariable, ByVal docID As Integer, ByVal docShopID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        Dim dtResult As DataTable
        Dim i As Integer
        Dim dclLastAmount, dclDisplayLastAmount, dclLastNetPrice, dclLastTax, refNetPrice As Decimal
        Dim dLastDate As Date
        Dim rResult() As DataRow
        Dim dtLastAmount, dtAllUnitRatio As DataTable
        Dim dclTotalPriceBeforeDiscount, dclDiscount As Decimal
        Dim remark As String = ""
        Dim matTemp As Decimal
        Dim testTemp As Decimal
        Dim testApi As Decimal
        dtResult = DocumentSQL.GetDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, docID, docShopID)

        dtLastAmount = New DataTable
        dtAllUnitRatio = New DataTable
        docData.DocDetailList = New List(Of DocumentDetail_Data)

        If dtResult.Rows.Count > 0 Then
            For i = 0 To dtResult.Rows.Count - 1
                dclLastAmount = dtResult.Rows(i)("ProductAmount")
                dclDiscount = dtResult.Rows(i)("ProductPricePerUnit")
                dclTotalPriceBeforeDiscount = dclLastAmount * dclDiscount
                dclDiscount = ((dclTotalPriceBeforeDiscount * dtResult.Rows(i)("ProductDiscount")) / 100) + dtResult.Rows(i)("ProductDiscountAmount")
                If dtLastAmount.Rows.Count = 0 Then
                    dclLastAmount = 0
                    dclDisplayLastAmount = 0
                    dclLastNetPrice = 0
                    dclLastTax = 0
                    dLastDate = Date.MinValue
                Else
                    rResult = dtLastAmount.Select("MaterialID = " & dtResult.Rows(i)("ProductID"))
                    If rResult.Length = 0 Then
                        dclLastAmount = 0
                        dclDisplayLastAmount = 0
                        dclLastNetPrice = 0
                        dclLastTax = 0
                        dLastDate = Date.MinValue
                    Else
                        dclLastAmount = rResult(0)("UnitSmallAmount")
                        dclDisplayLastAmount = ConvertAmountFromSmallToLargeAmount(dclLastAmount, dtAllUnitRatio, dtResult.Rows(i)("ProductUnit"), dtResult.Rows(i)("UnitID"))
                        dclLastNetPrice = rResult(0)("MaterialNetPrice")
                        dclLastTax = rResult(0)("MaterialVAT")
                        If Not IsDBNull(rResult(0)("MaterialLastDate")) Then
                            dLastDate = rResult(0)("MaterialLastDate")
                        Else
                            dLastDate = Date.MinValue
                        End If
                    End If
                End If
                If dtResult.Rows(i)("ProductTaxType") = 1 Then
                    refNetPrice = dtResult.Rows(i)("ProductTax") + dtResult.Rows(i)("ProductNetPrice")
                Else
                    refNetPrice = dtResult.Rows(i)("ProductNetPrice")
                End If
                If Not IsDBNull(dtResult.Rows(i)("ExtraText1")) Then
                    remark = dtResult.Rows(i)("ExtraText1")
                Else
                    remark = ""
                End If
                matTemp = dtResult.Rows(i)("ExtraValue1")
                testTemp = dtResult.Rows(i)("ExtraValue2")
                testApi = dtResult.Rows(i)("ExtraValue3")
                DocumentDetail_Data.AddOrUpdateDocDetailData(docData.DocDetailList, -1,
                    dtResult.Rows(i)("DocumentID"), dtResult.Rows(i)("ShopID"), dtResult.Rows(i)("DocDetailID"),
                    0, dtResult.Rows(i)("ProductID"), dtResult.Rows(i)("MaterialCode"),
                    dtResult.Rows(i)("MaterialName"), dtResult.Rows(i)("ProductTaxType"), dtResult.Rows(i)("UnitName"),
                    dtResult.Rows(i)("UnitID"), dtResult.Rows(i)("ProductUnit"), dtResult.Rows(i)("ProductAmount"), dtResult.Rows(i)("ProductPricePerUnit"),
                    dclTotalPriceBeforeDiscount, dclDiscount, dtResult.Rows(i)("ProductDiscount"),
                    dtResult.Rows(i)("ProductDiscountAmount"), dtResult.Rows(i)("ProductTax"), dtResult.Rows(i)("ProductNetPrice"),
                    dtResult.Rows(i)("UnitSmallAmount"), dtResult.Rows(i)("RequestSmallAmount"), dtResult.Rows(i)("PrepareSmallAmount"),
                    dtResult.Rows(i)("TransferSmallAmount"), dtResult.Rows(i)("ROSmallAmount"), dtResult.Rows(i)("DefaultInCompare"),
                    dclLastAmount, dclDisplayLastAmount, dclLastNetPrice,
                    dclLastTax, dLastDate, refNetPrice, dtResult.Rows(i)("ReferenceProductTax"), dtResult(i)("StockAmount"), dtResult.Rows(i)("DiffStockAmount"),
                    remark, "", matTemp, testTemp, testApi)
            Next i

        End If

        resultText = ""
        Return True
    End Function

    Friend Function LoadDocumentDetail(ByVal globalVariable As GlobalVariable, ByVal docID As Integer, ByVal docShopID As Integer, ByRef docData As DocumentPTT_Data, ByRef resultText As String) As Boolean
        Dim dtResult As DataTable
        Dim i As Integer
        Dim dclLastAmount, dclDisplayLastAmount, dclLastNetPrice, dclLastTax, refNetPrice As Decimal
        Dim dLastDate As Date
        Dim rResult() As DataRow
        Dim dtLastAmount, dtAllUnitRatio As DataTable
        Dim dclTotalPriceBeforeDiscount, dclDiscount As Decimal
        Dim remark As String = ""
        Dim api60F As String = ""
        Dim matTemp As Decimal
        Dim testTemp As Decimal
        Dim testApi As Decimal

        dtResult = DocumentSQL.GetDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, docID, docShopID)

        dtLastAmount = New DataTable
        dtAllUnitRatio = New DataTable
        docData.DocDetailList = New List(Of DocumentDetail_Data)

        If dtResult.Rows.Count > 0 Then
            For i = 0 To dtResult.Rows.Count - 1
                dclLastAmount = dtResult.Rows(i)("ProductAmount")
                dclDiscount = dtResult.Rows(i)("ProductPricePerUnit")
                dclTotalPriceBeforeDiscount = dclLastAmount * dclDiscount
                dclDiscount = ((dclTotalPriceBeforeDiscount * dtResult.Rows(i)("ProductDiscount")) / 100) + dtResult.Rows(i)("ProductDiscountAmount")
                If dtLastAmount.Rows.Count = 0 Then
                    dclLastAmount = 0
                    dclDisplayLastAmount = 0
                    dclLastNetPrice = 0
                    dclLastTax = 0
                    dLastDate = Date.MinValue
                Else
                    rResult = dtLastAmount.Select("MaterialID = " & dtResult.Rows(i)("ProductID"))
                    If rResult.Length = 0 Then
                        dclLastAmount = 0
                        dclDisplayLastAmount = 0
                        dclLastNetPrice = 0
                        dclLastTax = 0
                        dLastDate = Date.MinValue
                    Else
                        dclLastAmount = rResult(0)("UnitSmallAmount")
                        dclDisplayLastAmount = ConvertAmountFromSmallToLargeAmount(dclLastAmount, dtAllUnitRatio, dtResult.Rows(i)("ProductUnit"), dtResult.Rows(i)("UnitID"))
                        dclLastNetPrice = rResult(0)("MaterialNetPrice")
                        dclLastTax = rResult(0)("MaterialVAT")
                        If Not IsDBNull(rResult(0)("MaterialLastDate")) Then
                            dLastDate = rResult(0)("MaterialLastDate")
                        Else
                            dLastDate = Date.MinValue
                        End If
                    End If
                End If
                If dtResult.Rows(i)("ProductTaxType") = 1 Then
                    refNetPrice = dtResult.Rows(i)("ProductTax") + dtResult.Rows(i)("ProductNetPrice")
                Else
                    refNetPrice = dtResult.Rows(i)("ProductNetPrice")
                End If
                If Not IsDBNull(dtResult.Rows(i)("ExtraText1")) Then
                    remark = dtResult.Rows(i)("ExtraText1")
                Else
                    remark = ""
                End If
                If Not IsDBNull(dtResult.Rows(i)("ExtraText2")) Then
                    api60F = dtResult.Rows(i)("ExtraText2")
                Else
                    api60F = ""
                End If
                matTemp = dtResult.Rows(i)("ExtraValue1")
                testTemp = dtResult.Rows(i)("ExtraValue2")
                testApi = dtResult.Rows(i)("ExtraValue3")

                DocumentDetail_Data.AddOrUpdateDocDetailData(docData.DocDetailList, -1,
                    dtResult.Rows(i)("DocumentID"), dtResult.Rows(i)("ShopID"), dtResult.Rows(i)("DocDetailID"),
                    0, dtResult.Rows(i)("ProductID"), dtResult.Rows(i)("MaterialCode"),
                    dtResult.Rows(i)("MaterialName"), dtResult.Rows(i)("ProductTaxType"), dtResult.Rows(i)("UnitName"),
                    dtResult.Rows(i)("UnitID"), dtResult.Rows(i)("ProductUnit"), dtResult.Rows(i)("ProductAmount"), dtResult.Rows(i)("ProductPricePerUnit"),
                    dclTotalPriceBeforeDiscount, dclDiscount, dtResult.Rows(i)("ProductDiscount"),
                    dtResult.Rows(i)("ProductDiscountAmount"), dtResult.Rows(i)("ProductTax"), dtResult.Rows(i)("ProductNetPrice"),
                    dtResult.Rows(i)("UnitSmallAmount"), dtResult.Rows(i)("RequestSmallAmount"), dtResult.Rows(i)("PrepareSmallAmount"),
                    dtResult.Rows(i)("TransferSmallAmount"), dtResult.Rows(i)("ROSmallAmount"), dtResult.Rows(i)("DefaultInCompare"),
                    dclLastAmount, dclDisplayLastAmount, dclLastNetPrice,
                    dclLastTax, dLastDate, refNetPrice, dtResult.Rows(i)("ReferenceProductTax"), dtResult(i)("StockAmount"), dtResult.Rows(i)("DiffStockAmount"),
                    remark, api60F, matTemp, testTemp, testApi)
            Next i
        End If

        resultText = ""
        Return True
    End Function

    Friend Function StatusDocument(ByVal globalVariable As GlobalVariable, ByRef statusList As List(Of Status_Data), ByRef resultText As String) As Boolean
        Dim dt As New DataTable
        Dim ptData As Status_Data
        Try
            dt = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
            If dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    ptData = New Status_Data
                    ptData.StatusId = dt.Rows(i)("StatusId")
                    ptData.StatusName = dt.Rows(i)("StatusName")
                    statusList.Add(ptData)
                Next
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return True
            End If
        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function CheckValidDocumentForCancelDocument(ByVal globalVariable As GlobalVariable, ByVal documentStatus As Integer, ByVal inventoryId As Integer,
                                                        ByVal documentDate As Date, ByRef resultText As String) As Boolean
        Select Case documentStatus
            Case globalVariable.DOCUMENTSTATUS_APPROVE
                Dim lastEndDay As Date
                lastEndDay = DocumentModule.GetLastEndDayDocumentDate(globalVariable, inventoryId, New Date)
                If lastEndDay <> Date.MinValue Then
                    If documentDate <= lastEndDay Then
                        resultText = "ไม่สามารถทำการยกเลิกข้อมูลเอกสารนี้ได้ เนื่องจากวันที่ของเอกสารอยู่ก่อนหน้าการทำปรับสต๊อกของคลังสินค้านี้ไปแล้ว"
                        Return False
                    End If
                End If
            Case Else
        End Select
        Return True
    End Function

    Friend Function UpdateMaterialDefaultPrice(ByVal globalVariable As GlobalVariable, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                               ByVal documentShopID As Integer, ByVal defaultForVendorID As Integer) As Integer


        Dim dtResult As DataTable
        Dim i As Integer
        Dim strIn As String
        Dim curMaterialID As Integer
        Dim dclDefaultPrice As Decimal

        dtResult = DocumentSQL.GetDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, documentID, documentShopID)
        strIn = ""
        For i = 0 To dtResult.Rows.Count - 1
            strIn &= dtResult.Rows(i)("MaterialID") & ", "
        Next i
        If strIn <> "" Then
            DocumentSQL.DeleteMaterialInTableDefaultPrice(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, documentShopID, defaultForVendorID, strIn)
        End If
        For i = 0 To dtResult.Rows.Count - 1
            If dtResult.Rows(i)("UnitSmallAmount") = 0 Then
                dclDefaultPrice = 0
            Else
                dclDefaultPrice = (dtResult.Rows(i)("ProductPricePerUnit") * dtResult.Rows(i)("ProductAmount")) / dtResult.Rows(i)("UnitSmallAmount")
            End If
            If curMaterialID <> dtResult.Rows(i)("MaterialID") Then
                DocumentSQL.InsertMaterialDefaultPrice(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, documentShopID, defaultForVendorID,
                                                      dtResult.Rows(i)("UnitID"), dclDefaultPrice, dtResult.Rows(i)("UnitSmallAmount"),
                                                      dtResult.Rows(i)("UnitSmallID"), 1, dtResult.Rows(i)("MaterialID"))
                curMaterialID = dtResult.Rows(i)("MaterialID")
            Else
                DocumentSQL.UpdateMaterialDefaultPrice(globalVariable.DocDBUtil, globalVariable.DocConn, objTrans, documentShopID, defaultForVendorID,
                                                       dtResult.Rows(i)("UnitID"), dclDefaultPrice, dtResult.Rows(i)("UnitSmallAmount"),
                                                       dtResult.Rows(i)("UnitSmallID"), 1, dtResult.Rows(i)("MaterialID"))
            End If
        Next i
        Return curMaterialID
    End Function

End Module
