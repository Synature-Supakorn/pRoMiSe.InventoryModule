Imports pRoMiSe.Utilitys.Utilitys
Imports System.Data.SqlClient

Module ReceiveFromTrnasferOrderModule

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
            DocumentSQL.InsertDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, selDocDetailID, materialID, addAmount,
                                             selUnitSmallID, selUnitID, selUnitName, selUnitSmallAmount, selMaterialCode, selMaterialName,
                                             selMaterialSupplierCode, selMaterialSupplierName)

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "ReceiveFromTrnasferOrderModule", "AddDocDetail", "99", ex.ToString)
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
        updateDate = Now
        strUpdateDate = FormatDateTime(updateDate)
        dbTrans = globalVariable.DocConn.BeginTransaction
        Try
            DocumentSQL.ApproveDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId, globalVariable.StaffID, strUpdateDate)
            MaterialSQL.AutoAddDailyStockMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
            MaterialSQL.AutoAddMonthlyStockMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)
            MaterialSQL.AutoAddWeeklyStockMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopId)

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "ReceiveFromTrnasferOrderModule", "ApproveDocument", "99", ex.ToString)
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
            If DocumentSQL.DocumentIsAlreadyReferTo(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, dtDocument.Rows(0)("DocumentIDRef"), dtDocument.Rows(0)("DocumentIDRefShopID"), dtDocument.Rows(0)("DocumentTypeID"), documentId, documentShopId) = False Then
                DocumentSQL.UpdateDocumentStatus(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, dtDocument.Rows(0)("DocumentIDRef"), dtDocument.Rows(0)("DocumentIDRefShopID"), globalVariable.DOCUMENTSTATUS_APPROVE, strUpdateDate, globalVariable.StaffID)
            End If

            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.Message
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "ReceiveFromTrnasferOrderModule", "CancelDocument", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

   
    Friend Function SearchDocument(ByVal globalVariable As GlobalVariable, ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date,
                                   ByVal searchInventoryID As Integer, ByVal searchFromInventoryId As Integer, ByRef docList As List(Of SearchDocumentResult_Data),
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
            dtResult = DocumentSQL.SearchDocument(globalVariable.DocDBUtil, globalVariable.DocConn, globalVariable.DOCUMENTTYPE_ROTRANSFER, strFromDate, strToDate, documentStatus, searchInventoryID, searchFromInventoryId, globalVariable.DocLangID)
            dtDocStatus = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
            If dtResult.Rows.Count > 0 Then
                docList = DocumentModule.InsertResultDataIntoList(globalVariable, globalVariable.DOCUMENTTYPE_ROTRANSFER, dtResult, dtDocStatus)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "ReceiveFromTrnasferOrderModule", "SearchDocument", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function SearchTransferOrderDocument(ByVal globalVariable As GlobalVariable, ByVal startDate As Date, ByVal endDate As Date, ByVal searchInventoryID As Integer,
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
        Dim isTransferOrder As Boolean = True
        
        Try
            dtResult = DocumentSQL.SearchDocumentForCreateNewDocument(globalVariable.DocDBUtil, globalVariable.DocConn, globalVariable.DOCUMENTTYPE_TRANSFER, strFromDate, strToDate, globalVariable.DOCUMENTSTATUS_APPROVE, searchInventoryID, isTransferOrder, globalVariable.DocLangID)
            dtDocStatus = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
            If dtResult.Rows.Count > 0 Then
                docList = DocumentModule.InsertResultDataIntoList(globalVariable, globalVariable.DOCUMENTTYPE_TRANSFER, dtResult, dtDocStatus)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "ReceiveFromTrnasferOrderModule", "SearchTransferOrderDocument", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function SaveDocumentDataIntoDB(ByVal globalVariable As GlobalVariable, ByVal documentId As Integer, ByVal inventoryID As Integer,
                                           ByVal documenttypeId As Integer, ByVal documentDate As Date, ByVal dueDate As DateTime, ByVal documentNote As String,
                                           ByVal invoiceReference As String, ByRef resultText As String) As Boolean

        Dim strDocDate, strDueDate, strUpdateDate As String
        Dim updateDate As DateTime
        Dim newSend As Integer = 0
        Dim dbTrans As SqlTransaction

        strDocDate = FormatDate(documentDate)
        If dueDate <> Date.MinValue Then
            strDueDate = FormatDateTime(dueDate)
        Else
            strDueDate = "NULL"
        End If

        updateDate = Now
        strUpdateDate = FormatDateTime(updateDate)
        newSend = 0

        dbTrans = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
        Try
            DocumentSQL.UpdateDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, inventoryID, inventoryID, strDocDate, strDueDate, Trim(documentNote), Trim(invoiceReference), strUpdateDate, globalVariable.StaffID)
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "ReceiveFromTrnasferOrderModule", "SaveDocumentDataIntoDB", "99", ex.ToString)
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
            DocumentSQL.UpdateDocumentDetail(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentId, documentShopID, docDetailId, materialID, addAmount, selUnitSmallID, selUnitID, selUnitName, selUnitSmallAmount, selMaterialCode, selMaterialName, selMaterialSupplierCode, selMaterialSupplierName)
            dbTrans.Commit()
        Catch ex As Exception
            resultText = ex.ToString
            dbTrans.Rollback()
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "ReceiveFromTrnasferOrderModule", "UpdateDocDetail", "99", ex.ToString)
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
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "ReceiveFromTrnasferOrderModule", "DeleteDocDetail", "99", ex.ToString)
            Return False
        End Try

        resultText = ""
        Return True
    End Function

End Module
