Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient
Imports pRoMiSe.Utilitys.Utilitys

Public Class AdjustOrderController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function ApproveDocument(ByVal documentTypeId As Integer, ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer,
                                    ByVal documentDate As Date, ByVal documentNote As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        Dim dtDocType As New DataTable
        Dim stockEnough As New List(Of MaterialNotEnoughStock_Data)
        If AdjustOrderModule.SaveDocumentDataIntoDB(globalVariable, documentTypeId, documentID, documentShopID, inventoryID, documentDate, documentNote, resultText) = False Then
            Return False
        End If
        dtDocType = DocumentSQL.GetDocumentTypeRedue(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryID, documentTypeId)
        If dtDocType.Rows.Count > 0 Then
            If CheckMaterialInStockEnoughForTransferDocument(documentID, documentShopID, documentDate, stockEnough, resultText) = True Then
                If AdjustOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
                    Return False
                End If
            Else
                Return False
            End If
        Else
            If AdjustOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
                Return False
            End If
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AddMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal materialID As Integer, ByVal addAmount As Decimal,
                                                         ByVal materialUnitLargeID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If AdjustOrderModule.AddDocDetail(globalVariable, documentID, documentShopID, materialID, addAmount, materialUnitLargeID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function CreateNewDocument(ByVal documentTypeId As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date, ByVal documentNote As String,
                                      ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If DocumentModule.CreateNewDocument(globalVariable, documentTypeId, inventoryID, inventoryID, Date.Now, docData, resultText) = False Then
            Return False
        End If
        If AdjustOrderModule.SaveDocumentDataIntoDB(globalVariable, documentTypeId, docData.DocumentID, docData.DocumentShopID, inventoryID, Date.Now, documentNote, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, docData.DocumentID, docData.DocumentShopID, docData, resultText)
    End Function

    Public Function CancelDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If AdjustOrderModule.CancelDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function CheckMaterialInStockEnoughForTransferDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal checkStockDate As Date,
                                                                 ByRef notEnoughStockDetail As List(Of MaterialNotEnoughStock_Data), ByRef resultText As String) As Boolean

        Dim dtNotEnoughStock As New DataTable
        Dim bolAutoCreateDRO As Boolean
        Dim isFromMatchingTable As Boolean = False
        Dim listOnlyEnoughStock As Boolean = True
        notEnoughStockDetail = New List(Of MaterialNotEnoughStock_Data)
        If DocumentSQL.CheckMaterialInStockAndCalculateAveragePricePerUnitForTransfer(globalVariable.DocDBUtil, globalVariable.DocConn, documentID, documentShopID, FormatDate(checkStockDate), documentShopID, globalVariable.StaffID, isFromMatchingTable, bolAutoCreateDRO, listOnlyEnoughStock, dtNotEnoughStock) = False Then
            notEnoughStockDetail = TransferOrderModule.InsertNotEnoughStockMaterialIntoList(dtNotEnoughStock)
            resultText = globalVariable.MESSAGE_NOTENOUGHSTOCK
            Return False
        Else
            resultText = ""
            Return True
        End If
    End Function

    Public Function DeleteMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal strDocDetailId As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If AdjustOrderModule.DeleteDocDetail(globalVariable, documentID, documentShopID, strDocDetailId, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function GetDocumentType(ByRef docTypeList As List(Of AddReduceDocumentType_Data), ByRef resultText As String) As Boolean
        Try
            docTypeList = AddReduceDocumentType_Data.ListDocTypeAddReduce()
        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function DocumentTypeAddStock(ByVal docData As Document_Data, ByRef addDocumentTypeList As List(Of AddReduceDocumentType_Data),
    ByRef resultText As String) As Boolean

        Return AdjustOrderModule.GetADocumentTypeAddRedueStock(globalVariable, AddReduceDocumentType_Data.MOVEMENT_ADD, addDocumentTypeList, resultText)
    End Function

    Public Function DocumentTypeRedueStock(ByVal docData As Document_Data, ByRef reduceDocumentTypeList As List(Of AddReduceDocumentType_Data),
    ByRef resultText As String) As Boolean
        Return AdjustOrderModule.GetADocumentTypeAddRedueStock(globalVariable, AddReduceDocumentType_Data.MOVEMENT_REDUCE, reduceDocumentTypeList, resultText)
    End Function

    Public Function SearchDocument(ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date, ByVal documentTypeId As Integer,
                                   ByVal inventoryId As Integer, ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean
        Return AdjustOrderModule.SearchDocument(globalVariable, documentStatus, startDate, endDate, documentTypeId, inventoryId, docList, resultText)
    End Function

    Public Function SaveDocument(ByVal documentTypeId As Integer, ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer,
                                 ByVal documentDate As Date, ByVal documentNote As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        If AdjustOrderModule.SaveDocumentDataIntoDB(globalVariable, documentTypeId, documentID, documentShopID, inventoryID, documentDate, documentNote, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, inventoryID, docData, resultText)
    End Function

    Public Function UpdateMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer, ByVal materialID As Integer,
                                              ByVal addAmount As Decimal, ByVal materialUnitLargeID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If AdjustOrderModule.UpdateDocDetail(globalVariable, documentID, documentShopID, docDetailId, materialID, addAmount, materialUnitLargeID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

End Class
