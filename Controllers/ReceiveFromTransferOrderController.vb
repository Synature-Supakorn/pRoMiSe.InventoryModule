Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient
Imports pRoMiSe.Utilitys.Utilitys

Public Class ReceiveFromTransferOrderController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer,  ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromTrnasferOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date,
                                    ByVal documentNote As String, ByVal documentRefNumber As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        If ReceiveFromTrnasferOrderModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, globalVariable.DOCUMENTTYPE_ROTRANSFER, documentDate, Date.MinValue, documentNote, docData.DocumentRefNumber, resultText) = False Then
            Return False
        End If
        If ReceiveFromTrnasferOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AddMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal materialID As Integer, ByVal addAmount As Decimal,
                                           ByVal materialUnitLargeID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromTrnasferOrderModule.AddDocDetail(globalVariable, documentID, documentShopID, materialID, addAmount, materialUnitLargeID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function CreateNewDocument(ByVal refDocumentID As Integer, ByVal refDocumentShopID As Integer, ByVal inventoryID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        Dim dtDoc As New DataTable
        Dim dtDocRef As New DataTable
        Dim documentId As Integer = 0
        Dim documentShopId As Integer = 0
        Dim deliveryCost As Decimal = 0

        dtDocRef = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, refDocumentID, refDocumentShopID, globalVariable.DocLangID)
        If dtDocRef.Rows.Count > 0 Then
            Select Case dtDocRef.Rows(0)("ReceiveBy")
                Case Is = globalVariable.DOCUMENTSTATUS_TEMP
                    If DocumentModule.CreateNewDocumentFromReferedDocument(globalVariable, globalVariable.DOCUMENTTYPE_ROTRANSFER, inventoryID, refDocumentID, refDocumentShopID, Date.Now, True, docData, resultText) = False Then
                        Return False
                    End If

                    If ReceiveFromPurchaseOrderModule.SaveDocumentDataIntoDB(globalVariable, docData.DocumentID, inventoryID, globalVariable.DOCUMENTTYPE_ROTRANSFER, Date.Now, docData.VendorID, docData.VendorGroupID, docData.DocumentNote, docData.DocumentRefNumber, deliveryCost, resultText) = False Then
                        Return False
                    End If
                    documentId = docData.DocumentID
                    documentShopId = docData.DocumentShopID
                Case Else
                    dtDoc = DocumentSQL.GetDocumentIDRef(globalVariable.DocDBUtil, globalVariable.DocConn, globalVariable.DOCUMENTTYPE_ROTRANSFER, refDocumentID, refDocumentShopID)
                    If dtDoc.Rows.Count > 0 Then
                        documentId = dtDoc.Rows(0)("DocumentId")
                        documentShopId = dtDoc.Rows(0)("ShopId")
                    End If
            End Select
        End If

        Return DocumentModule.LoadDocument(globalVariable, documentId, documentShopId, docData, resultText)
    End Function

    Public Function CancelDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromTrnasferOrderModule.CancelDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function DeleteMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal strDocDetailId As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromTrnasferOrderModule.DeleteDocDetail(globalVariable, documentID, documentShopID, strDocDetailId, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function SearchDocument(ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date, ByVal inventoryID As Integer, ByVal fromInventoryID As Integer, ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean
        Return ReceiveFromTrnasferOrderModule.SearchDocument(globalVariable, documentStatus, startDate, endDate, inventoryID, fromInventoryID, docList, resultText)
    End Function

    Public Function SearchTransferOrderDocument(ByVal startDate As Date, ByVal endDate As Date, ByVal inventoryID As Integer, ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean
        Return ReceiveFromTrnasferOrderModule.SearchTransferOrderDocument(globalVariable, startDate, endDate, inventoryID, docList, resultText)
    End Function

    Public Function SaveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date, ByVal documentNote As String, ByVal documentRefNumber As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromTrnasferOrderModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, globalVariable.DOCUMENTTYPE_ROTRANSFER, documentDate, Date.MinValue, documentNote, documentRefNumber, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, inventoryID, docData, resultText)
    End Function

    Public Function UpdateMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer, ByVal materialID As Integer,
                                              ByVal addAmount As Decimal, ByVal materialUnitLargeID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromTrnasferOrderModule.UpdateDocDetail(globalVariable, documentID, documentShopID, docDetailId, materialID, addAmount, materialUnitLargeID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

End Class
