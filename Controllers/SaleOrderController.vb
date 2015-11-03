Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient
Imports pRoMiSe.Utilitys.Utilitys
Imports Newtonsoft.Json

Public Class SaleOrderController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If AdjustOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function ApproveDocument(ByVal documentTypeId As Integer, ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer,
                                    ByVal documentDate As Date, ByVal documentNote As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        If AdjustOrderModule.SaveDocumentDataIntoDB(globalVariable, documentTypeId, documentID, documentShopID, inventoryID, documentDate, documentNote, resultText) = False Then
            Return False
        End If
        If AdjustOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AddMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal materialID As Integer, ByVal addAmount As Decimal,
                                                         ByVal materialUnitLargeID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        Dim dtResult As New DataTable
        dtResult = DocumentSQL.GetDocumentDetailByMaterialID(globalVariable.DocDBUtil, globalVariable.DocConn, documentID, documentShopID, materialID)
        If dtResult.Rows.Count > 0 Then
            addAmount += dtResult.Rows(0)("ProductAmount")
            If AdjustOrderModule.UpdateDocDetail(globalVariable, documentID, documentShopID, dtResult.Rows(0)("DocDetailID"), materialID, addAmount, materialUnitLargeID, resultText) = False Then
                Return False
            End If
        Else
            If AdjustOrderModule.AddDocDetail(globalVariable, documentID, documentShopID, materialID, addAmount, materialUnitLargeID, resultText) = False Then
                Return False
            End If
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function CreateNewDocument(ByVal documentTypeId As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date, ByVal documentNote As String,
                                      ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        Dim dtResult As New DataTable
        Dim documentId, documentShopId As Integer

        dtResult = DocumentSQL.GetSaleDocument(globalVariable.DocDBUtil, globalVariable.DocConn, FormatDate(documentDate), inventoryID, globalVariable.DocLangID, documentTypeId)
        If dtResult.Rows.Count > 0 Then
            documentId = dtResult.Rows(0)("documentId")
            documentShopId = dtResult.Rows(0)("ShopId")
            Return DocumentModule.LoadDocument(globalVariable, documentId, documentShopId, docData, resultText)
        End If
        If DocumentModule.CreateNewDocument(globalVariable, documentTypeId, inventoryID, inventoryID, documentDate, docData, resultText) = False Then
            Return False
        End If
        
        If AdjustOrderModule.SaveDocumentDataIntoDB(globalVariable, docData.DocumentID, docData.DocumentShopID, inventoryID, documentTypeId, documentDate, documentNote, resultText) = False Then
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

    Public Function DeleteMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal strDocDetailId As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If AdjustOrderModule.DeleteDocDetail(globalVariable, documentID, documentShopID, strDocDetailId, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
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
