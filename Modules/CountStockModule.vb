Imports pRoMiSe.Utilitys.Utilitys
Imports System.Data.SqlClient
Imports System.Text

Module CountStockModule

    Friend Function AddDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopID As Integer, ByVal materialID As Integer,
                             ByVal addAmount As Decimal, ByVal materialUnitLargeID As Integer, ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        Dim dtMaterialUnit As DataTable
        Dim selDocDetailID As Integer
        Dim selUnitSmallAmount As Decimal
        Dim selUnitSmallID, selUnitID As Integer
        Dim selUnitName As String
        Dim selMaterialCode As String = ""
        Dim selMaterialName As String = ""
        Dim selMaterialSupplierCode As String = ""
        Dim selMaterialSupplierName As String = ""
        Dim stockAmount As Decimal = 0
        Dim diffStockAmount As Decimal = 0
        Dim isAddRedueStock As Integer = 0
        Dim dtStock As New DataTable
        Dim startDate As Date
        Dim endDate As Date

        startDate = New Date(Now.Year, Now.Month, 1)
        endDate = Date.Now

        dtMaterialUnit = MaterialSQL.GetMaterialDetailAndUnitRatio(globalVariable.DocDBUtil, globalVariable.DocConn, materialID, materialUnitLargeID, False)
        If dtMaterialUnit.Rows.Count = 0 Then
            resultText = globalVariable.MESSAGE_MATERIALNOTFOUND
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

        dtStock = DocumentSQL.GetCurrentStock(globalVariable.DocDBUtil, globalVariable.DocConn, startDate, endDate, documentShopID, materialID)
        If dtStock.Rows.Count > 0 Then
            stockAmount = dtStock.Rows(0)("Qty")
        Else
            stockAmount = 0
        End If
        Select Case stockAmount
            Case Is = 0
                diffStockAmount = addAmount
                isAddRedueStock = 1
            Case Else
                If stockAmount = addAmount Then
                    diffStockAmount = 0
                    isAddRedueStock = 0
                Else
                    diffStockAmount = (stockAmount - addAmount)
                    If diffStockAmount > 0 Then
                        isAddRedueStock = 2
                    Else
                        isAddRedueStock = 1
                        diffStockAmount = Math.Abs(diffStockAmount)
                    End If
                End If
        End Select
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            selDocDetailID = DocumentSQL.GetMaxDocDetailID(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)
            DocumentSQL.InsertDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, selDocDetailID, materialID, addAmount,
                                             selUnitSmallID, selUnitID, selUnitName, selUnitSmallAmount, selMaterialCode, selMaterialName,
                                             selMaterialSupplierCode, selMaterialSupplierName, stockAmount, diffStockAmount, isAddRedueStock)
            If selDocDetailID = 1 Then
                DocumentSQL.UpdateStockAtDateTime(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID)
            End If
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "CountStockModule", "AddDocDetail", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function AutoAddDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopID As Integer, ByVal tableName As String, ByRef resultText As String) As Boolean
        Dim dtResult As New DataTable
        Dim ck As Boolean = True
        dtResult = MaterialSQL.GetMaterialFromCountStockSetting(globalVariable.DocDBUtil, globalVariable.DocConn, tableName)
        If dtResult.Rows.Count > 0 Then
            For i As Integer = 0 To dtResult.Rows.Count - 1
                If AddDocDetail(globalVariable, documentId, documentShopID, dtResult.Rows(i)("MaterialId"), 0, dtResult.Rows(i)("UnitID"), resultText) = False Then
                    ck = False
                    Exit For
                End If
            Next
        End If
        Return ck
    End Function

    Friend Function ApproveDocument(ByVal globalVariable As GlobalVariable, ByVal documentTypeId As Integer, ByVal documentId As Integer, ByVal documentShopId As Integer, ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        Dim updateDate As DateTime
        Dim strUpdateDate As String
        Dim strDocDate As String
        Dim documentDate As Date
        Dim vendorId As Integer = 0
        Dim dtResult As New DataTable
        Dim dtDocDetail As New DataTable
        Dim newDocId As Integer
        Dim newDocumentNumber As String = ""
        Dim expression As String = ""
        Dim foundRows() As DataRow
        Dim addStock As Integer = 1
        Dim redueStock As Integer = 2

        updateDate = Now
        strUpdateDate = FormatDateTime(updateDate)
        dtResult = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId, globalVariable.DocLangID)
        dtDocDetail = DocumentSQL.GetDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId)

        If dtDocDetail.Rows.Count = 0 Then
            resultText = globalVariable.MESSAGE_MATERIALNOTFOUND
            Return False
        End If
        documentDate = dtResult.Rows(0)("documentDate")
        strDocDate = FormatDate(documentDate)

        dbTrans = globalVariable.DocConn.BeginTransaction
        Try
             Select Case documentTypeId
                Case Is = globalVariable.DOCUMENTTYPE_DAILYSTOCK
                    DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, globalVariable.StaffID, strUpdateDate)
                    expression = "IsAddRedueStock=1"
                    foundRows = dtDocDetail.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, documentShopId)
                        DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentShopId, 0, globalVariable.DOCUMENTTYPE_DAILYSTOCK_ADD, globalVariable.DOCUMENTSTATUS_APPROVE, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
                        newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, documentShopId, globalVariable.DOCUMENTTYPE_DAILYSTOCK_ADD, documentDate, documentDate.Year, documentDate.Month)
                        DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.DOCUMENTTYPE_DAILYSTOCK_ADD, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
                        DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.StaffID, strUpdateDate)
                        DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentId, documentShopId)
                        DocumentSQL.InsertAddRedueDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, dtResult.Rows(0)("documentid"), dtResult.Rows(0)("shopID"), addStock)
                    End If
                    expression = "IsAddRedueStock=2"
                    foundRows = dtDocDetail.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, documentShopId)
                        DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentShopId, 0, globalVariable.DOCUMENTTYPE_DAILYSTOCK_REDUCE, globalVariable.DOCUMENTSTATUS_APPROVE, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
                        newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, documentShopId, globalVariable.DOCUMENTTYPE_DAILYSTOCK_REDUCE, documentDate, documentDate.Year, documentDate.Month)
                        DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.DOCUMENTTYPE_DAILYSTOCK_REDUCE, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
                        DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.StaffID, strUpdateDate)
                        DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentId, documentShopId)
                        DocumentSQL.InsertAddRedueDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, dtResult.Rows(0)("documentid"), dtResult.Rows(0)("shopID"), redueStock)
                    End If
                Case Is = globalVariable.DOCUMENTTYPE_WEEKLYSTOCK
                    DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, globalVariable.StaffID, strUpdateDate)
                    expression = "IsAddRedueStock=1"
                    foundRows = dtDocDetail.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, documentShopId)
                        DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentShopId, 0, globalVariable.DOCUMENTTYPE_WEEKLYSTOCK_ADD, globalVariable.DOCUMENTSTATUS_APPROVE, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
                        newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, documentShopId, globalVariable.DOCUMENTTYPE_WEEKLYSTOCK_ADD, documentDate, documentDate.Year, documentDate.Month)
                        DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.DOCUMENTTYPE_WEEKLYSTOCK_ADD, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
                        DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.StaffID, strUpdateDate)
                        DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentId, documentShopId)
                        DocumentSQL.InsertAddRedueDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, dtResult.Rows(0)("documentid"), dtResult.Rows(0)("shopID"), addStock)
                    End If
                    expression = "IsAddRedueStock=2"
                    foundRows = dtDocDetail.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, documentShopId)
                        DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentShopId, 0, globalVariable.DOCUMENTTYPE_WEEKLYSTOCK_REDUCE, globalVariable.DOCUMENTSTATUS_APPROVE, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
                        newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, documentShopId, globalVariable.DOCUMENTTYPE_WEEKLYSTOCK_REDUCE, documentDate, documentDate.Year, documentDate.Month)
                        DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.DOCUMENTTYPE_WEEKLYSTOCK_REDUCE, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
                        DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.StaffID, strUpdateDate)
                        DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentId, documentShopId)
                        DocumentSQL.InsertAddRedueDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, dtResult.Rows(0)("documentid"), dtResult.Rows(0)("shopID"), redueStock)
                    End If
                Case Is = globalVariable.DOCUMENTTYPE_MONTHLYSTOCK
                    Dim dt As Date = documentDate
                    Dim nextMonth As Date
                    Dim lastMonth As Date
                    nextMonth = New Date(dt.AddMonths(1).Year, dt.AddMonths(1).Month, 1)
                    lastMonth = New Date(dt.Year, dt.Month, System.DateTime.DaysInMonth(dt.Year, dt.Month))

                    'CREATE NEW DOCUMENTTYPE : 10 TRANSFERSTOCK
                    newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, documentShopId)
                    DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentShopId, 0, globalVariable.DOCUMENTTYPE_TRANSFERSTOCK, globalVariable.DOCUMENTSTATUS_APPROVE, FormatDate(nextMonth), "NULL", "", strUpdateDate, globalVariable.StaffID)
                    newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, documentShopId, globalVariable.DOCUMENTTYPE_TRANSFERSTOCK, documentDate, documentDate.Year, documentDate.Month)
                    DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.DOCUMENTTYPE_TRANSFERSTOCK, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
                    DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.StaffID, strUpdateDate)
                    DocumentSQL.AddDocDetailMonthlyStockTransfer(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, dt.Month, dt.Year)
                    DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentId, documentShopId)

                    strDocDate = FormatDate(lastMonth)
                    expression = "IsAddRedueStock=1"
                    foundRows = dtDocDetail.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, documentShopId)
                        DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentShopId, 0, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_ADD, globalVariable.DOCUMENTSTATUS_APPROVE, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
                        newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, documentShopId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_ADD, documentDate, documentDate.Year, documentDate.Month)
                        DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_ADD, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
                        DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.StaffID, strUpdateDate)
                        DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentId, documentShopId)
                        DocumentSQL.InsertAddRedueDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, dtResult.Rows(0)("documentid"), dtResult.Rows(0)("shopID"), addStock)
                    End If
                    expression = "IsAddRedueStock=2"
                    foundRows = dtDocDetail.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, documentShopId)
                        DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentShopId, 0, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_REDUCE, globalVariable.DOCUMENTSTATUS_APPROVE, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
                        newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, documentShopId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_REDUCE, documentDate, documentDate.Year, documentDate.Month)
                        DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_REDUCE, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
                        DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, globalVariable.StaffID, strUpdateDate)
                        DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, documentId, documentShopId)
                        DocumentSQL.InsertAddRedueDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, documentShopId, dtResult.Rows(0)("documentid"), dtResult.Rows(0)("shopID"), redueStock)
                    End If
                    'APPROVE COUNTMONTHLY STOCK
                    DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, globalVariable.StaffID, strUpdateDate)
            End Select
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "CountStockModule", "ApproveDocument", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function AutoTransferStock(ByVal globalVariable As GlobalVariable, ByVal documentDate As Date, ByVal inventoryId As Integer,
                                      ByRef monthlyDocumentId As Integer, ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        Dim strUpdateDate As String
        Dim strDocDate As String
        Dim vendorId As Integer = 0
        Dim dtResult As New DataTable
        Dim dtDocDetail As New DataTable
        Dim newDocId As Integer
        Dim newDocumentNumber As String = ""
        Dim expression As String = ""
        Dim foundRows() As DataRow
        Dim addStock As Integer = 1
        Dim redueStock As Integer = 2
        Dim monthlyDocId As Integer = 0

        strUpdateDate = FormatDateTime(Date.Now)
        dbTrans = globalVariable.DocConn.BeginTransaction
        Try
            Dim dt As Date = documentDate
            Dim nextMonth As Date
            Dim lastMonth As Date

            nextMonth = New Date(dt.AddMonths(1).Year, dt.AddMonths(1).Month, 1)
            lastMonth = New Date(dt.Year, dt.Month, System.DateTime.DaysInMonth(dt.Year, dt.Month))

            'CREATE NEW DOCUMENTTYPE : 7 MONTHLY STOCK
            monthlyDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, inventoryId)
            DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, monthlyDocId, inventoryId, inventoryId, 0, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, globalVariable.DOCUMENTSTATUS_WORKING, FormatDate(lastMonth), "NULL", "", strUpdateDate, globalVariable.StaffID)
            newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, inventoryId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, documentDate, documentDate.Year, documentDate.Month)
            DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, monthlyDocId, inventoryId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
            DocumentSQL.AddDocDetailMonthlyStockTransfer(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, monthlyDocId, inventoryId, dt.Month, dt.Year)
            DocumentSQL.RefreshMonthStockForAddRedueDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, monthlyDocId, inventoryId)

            'CREATE NEW DOCUMENTTYPE : 10 TRANSFERSTOCK
            newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, inventoryId)
            DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, inventoryId, 0, globalVariable.DOCUMENTTYPE_TRANSFERSTOCK, globalVariable.DOCUMENTSTATUS_APPROVE, FormatDate(nextMonth), "NULL", "", strUpdateDate, globalVariable.StaffID)
            newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, inventoryId, globalVariable.DOCUMENTTYPE_TRANSFERSTOCK, documentDate, documentDate.Year, documentDate.Month)
            DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, globalVariable.DOCUMENTTYPE_TRANSFERSTOCK, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
            DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, globalVariable.StaffID, strUpdateDate)
            DocumentSQL.AddDocDetailMonthlyStockTransfer(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, dt.Month, dt.Year)
            DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, monthlyDocId, inventoryId)

            dtDocDetail = DocumentSQL.GetDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, monthlyDocId, inventoryId)
            strDocDate = FormatDate(lastMonth)
            expression = "IsAddRedueStock=1"
            foundRows = dtDocDetail.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, inventoryId)
                DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, inventoryId, 0, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_ADD, globalVariable.DOCUMENTSTATUS_APPROVE, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
                newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, inventoryId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_ADD, documentDate, documentDate.Year, documentDate.Month)
                DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_ADD, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
                DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, globalVariable.StaffID, strUpdateDate)
                DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, monthlyDocId, inventoryId)
                DocumentSQL.InsertAddRedueDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, monthlyDocId, inventoryId, addStock)
            End If
            expression = "IsAddRedueStock=2"
            foundRows = dtDocDetail.Select(expression)
            If foundRows.GetUpperBound(0) >= 0 Then
                newDocId = DocumentModule.GetNewDocumentIDFromMaxDocumentID(globalVariable, dbTrans, inventoryId)
                DocumentSQL.CreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, inventoryId, 0, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_REDUCE, globalVariable.DOCUMENTSTATUS_APPROVE, strDocDate, "NULL", "", strUpdateDate, globalVariable.StaffID)
                newDocumentNumber = DocumentModule.GetAndUpdateDocumentNumber(globalVariable, dbTrans, inventoryId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_REDUCE, documentDate, documentDate.Year, documentDate.Month)
                DocumentSQL.InsertDocumentHeader(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK_REDUCE, documentDate.Month, documentDate.Year, newDocumentNumber, globalVariable.DocLangID)
                DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, globalVariable.StaffID, strUpdateDate)
                DocumentSQL.UpdateDocumentRef(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, monthlyDocId, inventoryId)
                DocumentSQL.InsertAddRedueDocDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, newDocId, inventoryId, monthlyDocId, inventoryId, redueStock)
            End If
            'APPROVE COUNTMONTHLY STOCK
            DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, monthlyDocId, inventoryId, globalVariable.StaffID, strUpdateDate)

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.Message
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "CountStockModule", "AutoTransferStock", "99", ex.ToString)
            Return False
        End Try
        monthlyDocumentId = monthlyDocId
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
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "CountStockModule", "CancelDocument", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function
   
    Friend Function CheckMaterialStockBelowZero(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopId As Integer,
                                                ByVal startDate As Date, ByVal endDate As Date, ByRef notEnoughStock As List(Of MaterialNotEnoughStock_Data),
                                                ByRef resultText As String) As Boolean

        Dim dtResult As New DataTable
        Dim dtDocDetail As New DataTable
        Dim strBD As New StringBuilder
        Dim strMaterialId As String = ""
        Dim dummyTable As String = GenerateGUID().ToString
        Dim dataZero As Integer = 0
        notEnoughStock = New List(Of MaterialNotEnoughStock_Data)
        dummyTable = dummyTable.Replace("-", "")

        dtDocDetail = DocumentSQL.GetDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopId)
        If dtDocDetail.Rows.Count > 0 Then
            For i As Integer = 0 To dtDocDetail.Rows.Count - 1
                strBD.Append("," & dtDocDetail.Rows(i)("ProductId"))
            Next
        End If
        If strBD.ToString <> "" Then
            strMaterialId = "0" & strBD.ToString
        Else
            strMaterialId = "0"
        End If
        dtResult = DocumentSQL.CheckMaterialStockBelowZero(globalVariable.DocDBUtil, globalVariable.DocConn, startDate, endDate, documentShopId, strMaterialId, dummyTable)
        If dtResult.Rows.Count > 0 Then
            For i As Integer = 0 To dtResult.Rows.Count - 1
                AddDocDetail(globalVariable, documentId, documentShopId, dtResult.Rows(i)("MaterialId"), 0, dtResult.Rows(i)("UnitSmallID"), resultText)
            Next
            For i As Integer = 0 To dtResult.Rows.Count - 1
                notEnoughStock.Add(MaterialNotEnoughStock_Data.NewMaterialNotEnoughStock(dtResult.Rows(i)("MaterialId"), dtResult.Rows(i)("MaterialCode"), dtResult.Rows(i)("MaterialName"), dataZero, dataZero, dataZero, dataZero, dtResult.Rows(i)("UnitSmallID")))
            Next
            resultText = globalVariable.MESSAGE_MATERIALBELOWZERO
            Return True
        Else
            resultText = ""
            Return False
        End If
    End Function

    Friend Function GetLastTransferStock(ByVal globalVariable As GlobalVariable, ByVal inventoryId As Integer) As String
        Dim dtResult As New DataTable
        Dim strDate As String = ""
        Dim tempDate As Date
        Dim InvC As System.Globalization.CultureInfo = System.Globalization.CultureInfo.InvariantCulture
        dtResult = DocumentSQL.GetLastTransferStock(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryId)
        If dtResult.Rows.Count > 0 Then
            tempDate = dtResult.Rows(0)("DocumentDate")
            strDate = tempDate.ToString("yyyy-MM-dd", InvC)
        End If
        Return strDate
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
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "CountStockModule", "DeleteDocDetail", "99", ex.ToString)
            Return False
        End Try

        resultText = ""
        Return True
    End Function
    
    Friend Function GetADocumentTypeAddRedueStock(ByVal clsVariable As GlobalVariable, ByVal movementInStock As Integer, ByRef addReduceDocList As List(Of AddReduceDocumentType_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtResult As DataTable
        Try
            dtResult = DocumentSQL.GetAddReduceDocumentType(clsVariable.DocDBUtil, clsVariable.DocConn, movementInStock, clsVariable.DocLangID,
                            clsVariable.DefaultDocShopID)
            addReduceDocList = New List(Of AddReduceDocumentType_Data)
            If dtResult.Rows.Count > 0 Then
                For i = 0 To dtResult.Rows.Count - 1
                    If IsDBNull(dtResult.Rows(i)("DocumentTypeHeader")) Then
                        dtResult.Rows(i)("DocumentTypeHeader") = ""
                    End If
                    If IsDBNull(dtResult.Rows(i)("DocumentTypeName")) Then
                        dtResult.Rows(i)("DocumentTypeName") = ""
                    End If
                    addReduceDocList.Add(AddReduceDocumentType_Data.NewAddReduceDocData(dtResult.Rows(i)("DocumentTypeID"),
                                    dtResult.Rows(i)("DocumentTypeHeader"), dtResult.Rows(i)("DocumentTypeName"), dtResult.Rows(i)("MovementInStock")))

                Next i
            Else
                resultText = GlobalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If

        Catch ex As Exception
            resultText = ex.ToString
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function InsertNotEnoughStockMaterialIntoList(ByVal dtNotEnoughStock As DataTable) As List(Of MaterialNotEnoughStock_Data)

        Dim notEnoughStockDetail As List(Of MaterialNotEnoughStock_Data)
        Dim curMaterialID As Integer
        Dim dataZero As Integer = 0
        notEnoughStockDetail = New List(Of MaterialNotEnoughStock_Data)
        For i = 0 To dtNotEnoughStock.Rows.Count - 1
            Do While (curMaterialID = dtNotEnoughStock.Rows(i)("MaterialID"))
                i += 1
                If i >= dtNotEnoughStock.Rows.Count Then
                    Exit For
                End If
            Loop
            curMaterialID = dtNotEnoughStock.Rows(i)("MaterialID")
            'Add To List
            notEnoughStockDetail.Add(MaterialNotEnoughStock_Data.NewMaterialNotEnoughStock(curMaterialID, dtNotEnoughStock.Rows(i)("MaterialCode"), dtNotEnoughStock.Rows(i)("MaterialName"), dataZero, dataZero, dataZero, dataZero, dataZero))
        Next i
        Return notEnoughStockDetail
    End Function
    
    Friend Function RefreshStockInDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopID As Integer,  ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        Dim selDocDetailID As Integer
        Dim CountAmount As Integer = 0
        Dim stockAmount As Decimal = 0
        Dim diffStockAmount As Decimal = 0
        Dim isAddRedueStock As Integer = 0
        Dim showAllMaterial As Integer = 0
        Dim dtStock As New DataTable
        Dim dtDocDetail As New DataTable
        Dim startDate As Date
        Dim endDate As Date
        Dim expression As String = ""
        Dim foundRows() As DataRow
        Dim strBD As New StringBuilder

        startDate = New Date(Now.Year, Now.Month, 1)
        endDate = Date.Now
        dtDocDetail = DocumentSQL.GetDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, documentId, documentShopID)
        If dtDocDetail.Rows.Count < 0 Then
            resultText = globalVariable.MESSAGE_MATERIALNOTFOUND
            Return False
        End If
        dtStock = DocumentSQL.GetCurrentStock(globalVariable.DocDBUtil, globalVariable.DocConn, startDate, endDate, documentShopID, showAllMaterial)
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try

            For i As Integer = 0 To dtDocDetail.Rows.Count - 1

                selDocDetailID = dtDocDetail.Rows(i)("DocdetialId")
                CountAmount = dtDocDetail.Rows(i)("ProductAmount")

                expression = "ProductId=" & dtDocDetail.Rows(i)("ProductId")
                foundRows = dtStock.Select(expression)
                If foundRows.GetUpperBound(0) >= 0 Then
                    stockAmount = foundRows(0)("Qty")
                Else
                    stockAmount = 0
                End If
                Select Case stockAmount
                    Case Is = 0
                        diffStockAmount = CountAmount
                        isAddRedueStock = 1
                    Case Else
                        diffStockAmount = (stockAmount - CountAmount)
                        If diffStockAmount > 0 Then
                            isAddRedueStock = 2
                        Else
                            isAddRedueStock = 1
                            diffStockAmount = Math.Abs(diffStockAmount)
                        End If
                End Select
                strBD.Append("UPDATE DocDetail Set StockAmount=" & stockAmount & ",DiffStockAmount=" & diffStockAmount & ",IsAddRedueStock=" & isAddRedueStock & " WHERE DocdetailId=" & selDocDetailID & " AND DocumentID=" & documentId & " AND ShopID=" & documentShopID & ";")
            Next
            globalVariable.DocDBUtil.sqlExecute(strBD.ToString, globalVariable.DocConn, dbTrans)
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "CountStockModule", "RefreshStockInDocDetail", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function SearchDocument(ByVal globalVariable As GlobalVariable, ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date,
                                   ByVal documentTypeId As Integer, ByVal inventoryId As Integer, ByRef docList As List(Of SearchDocumentResult_Data),
                                   ByRef resultText As String) As Boolean

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
            dtResult = DocumentSQL.SearchDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentTypeId, strFromDate, strToDate, documentStatus,
                                                  inventoryId, "", globalVariable.DocLangID)
            dtDocStatus = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
            If dtResult.Rows.Count > 0 Then
                docList = DocumentModule.InsertResultDataIntoList(globalVariable, documentTypeId, dtResult, dtDocStatus)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "CountStockModule", "SearchDocument", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function SaveDocumentDataIntoDB(ByVal globalVariable As GlobalVariable, ByVal documentTypeId As Integer, ByVal documentID As Integer, ByVal documentShopID As Integer,
                                           ByVal inventoryID As Integer, ByVal documentDate As Date, ByVal documentNote As String, ByRef resultText As String) As Boolean

        Dim strDocDate, strDueDate, strUpdateDate As String
        Dim updateDate As DateTime
        Dim newSend As Integer = 0
        Dim dbTrans As SqlTransaction
        Dim invoiceReference As String = ""

        strDocDate = FormatDate(documentDate)
        strDueDate = "NULL"
        updateDate = Now
        strUpdateDate = FormatDateTime(updateDate)
        newSend = 0

        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentSQL.UpdateDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentID, inventoryID, inventoryID, strDocDate, strDueDate, Trim(documentNote), Trim(invoiceReference), strUpdateDate, globalVariable.StaffID)
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "CountStockModule", "SaveDocumentDataIntoDB", "99", ex.ToString)
            Return False
        End Try

        resultText = ""
        Return True
    End Function

    Friend Function UpdateDocDetail(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer,
                                    ByVal materialID As Integer, ByVal addAmount As Decimal, ByVal materialUnitLargeID As Integer, ByRef resultText As String) As Boolean

        Dim dbTrans As SqlTransaction
        Dim dtMaterialUnit As DataTable
        Dim selUnitSmallAmount As Decimal
        Dim selUnitSmallID, selUnitID As Integer
        Dim selUnitName As String
        Dim selMaterialCode As String = ""
        Dim selMaterialName As String = ""
        Dim selMaterialSupplierCode As String = ""
        Dim selMaterialSupplierName As String = ""
        Dim stockAmount As Decimal = 0
        Dim diffStockAmount As Decimal = 0
        Dim isAddRedueStock As Integer = 0
        Dim dtStock As New DataTable
        Dim startDate As Date
        Dim endDate As Date

        startDate = New Date(Now.Year, Now.Month, 1)
        endDate = Date.Now

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
        dtStock = DocumentSQL.GetCurrentStock(globalVariable.DocDBUtil, globalVariable.DocConn, startDate, endDate, documentShopID, materialID)
        If dtStock.Rows.Count > 0 Then
            stockAmount = dtStock.Rows(0)("Qty")
        Else
            stockAmount = 0
        End If
        Select Case stockAmount
            Case Is = 0
                diffStockAmount = addAmount
                isAddRedueStock = 1
            Case Else
                diffStockAmount = (stockAmount - addAmount)
                If diffStockAmount > 0 Then
                    isAddRedueStock = 2
                Else
                    isAddRedueStock = 1
                    diffStockAmount = Math.Abs(diffStockAmount)
                End If
        End Select
        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentSQL.UpdateDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, docDetailId, materialID, addAmount,
                                             selUnitSmallID, selUnitID, selUnitName, selUnitSmallAmount, selMaterialCode, selMaterialName, selMaterialSupplierCode, selMaterialSupplierName,
                                             stockAmount, diffStockAmount, isAddRedueStock)

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "CountStockModule", "SaveDocumentDataIntoDB", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function
    
End Module
