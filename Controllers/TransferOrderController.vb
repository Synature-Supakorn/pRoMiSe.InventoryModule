Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient
Imports pRoMiSe.Utilitys.Utilitys

Public Class TransferOrderController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal toInventoryID As Integer,
                                    ByVal documentDate As Date, ByVal documentNote As String, ByVal deliveryDate As DateTime, ByVal invoiceReference As String,
                                    ByVal staffId As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        If TransferOrderModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, toInventoryID, globalVariable.DOCUMENTTYPE_TRANSFER, documentDate, deliveryDate, documentNote, invoiceReference, resultText) = False Then
            Return False
        End If
        Dim stockEnough As New List(Of MaterialNotEnoughStock_Data)
        If CheckMaterialInStockEnoughForTransferDocument(documentID, documentShopID, documentDate, stockEnough, resultText) = True Then
            If TransferOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
                Return False
            End If
        Else
            Dim msg As String = ""
            msg = resultText
            DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
            docData.MaterialNotEnoughStock = stockEnough
            resultText = msg
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AddMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal materialID As Integer, ByVal addAmount As Decimal,
                                                         ByVal materialUnitLargeID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If TransferOrderModule.AddDocDetail(globalVariable, documentID, documentShopID, materialID, addAmount, materialUnitLargeID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function CreateNewDocument(ByVal inventoryID As Integer, ByVal toInventoryID As Integer, ByVal documentDate As Date, ByVal documentNote As String,
                                      ByVal deliveryDate As DateTime, ByVal invoiceReference As String, ByRef docData As Document_Data,
                                      ByRef resultText As String) As Boolean
        If DocumentModule.CreateNewDocument(globalVariable, globalVariable.DOCUMENTTYPE_TRANSFER, inventoryID, toInventoryID, documentDate, docData, resultText) = False Then
            Return False
        End If
        If TransferOrderModule.SaveDocumentDataIntoDB(globalVariable, docData.DocumentID, docData.DocumentShopID, toInventoryID, globalVariable.DOCUMENTTYPE_TRANSFER, documentDate, deliveryDate, documentNote, invoiceReference, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, docData.DocumentID, docData.DocumentShopID, docData, resultText)
    End Function

    Public Function CancelDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If TransferOrderModule.CancelDocument(globalVariable, documentID, documentShopID, resultText) = False Then
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
        If TransferOrderModule.DeleteDocDetail(globalVariable, documentID, documentShopID, strDocDetailId, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function SearchDocument(ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date, ByVal inventoryID As Integer,
                                   ByVal toInventoryId As Integer, ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean
        Return TransferOrderModule.SearchDocument(globalVariable, documentStatus, startDate, endDate, inventoryID, toInventoryId, docList, resultText)
    End Function

    Public Function SaveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal toInventoryID As Integer,
                                 ByVal documentDate As Date, ByVal documentNote As String, ByVal deliveryDate As DateTime, ByVal invoiceReference As String,
                                 ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        If TransferOrderModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, toInventoryID, globalVariable.DOCUMENTTYPE_TRANSFER, documentDate,
                                                      deliveryDate, documentNote, invoiceReference, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, inventoryID, docData, resultText)
    End Function

    Public Function UpdateMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer, ByVal materialID As Integer, ByVal addAmount As Decimal,
                                                            ByVal materialUnitLargeID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If TransferOrderModule.UpdateDocDetail(globalVariable, documentID, documentShopID, docDetailId, materialID, addAmount, materialUnitLargeID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

End Class
