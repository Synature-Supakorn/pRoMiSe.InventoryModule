Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient
Imports pRoMiSe.Utilitys.Utilitys
Public Class ReceiveFromPurchaseOrderController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer,  ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromPurchaseOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date,
                                    ByVal vendorID As Integer, ByVal vendorGroupID As Integer, ByVal documentNote As String, ByVal invoiceReference As String,
                                    ByVal deliveryCost As Decimal, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromPurchaseOrderModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, globalVariable.DOCUMENTTYPE_ROPO, documentDate, vendorID, vendorGroupID, documentNote, invoiceReference, deliveryCost, resultText) = False Then
            Return False
        End If
        If ReceiveFromPurchaseOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AddMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByVal materialID As Integer,
                                           ByVal materailQty As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal, ByVal discountAmount As Decimal,
                                           ByVal discountPercent As Decimal, ByVal materialVATType As Integer, ByRef resultText As String) As Boolean
        If ReceiveFromPurchaseOrderModule.AddDocDetail(globalVariable, documentID, documentShopID, materialID, materailQty, materialUnitLargeID, pricePerUnit, discountAmount, discountPercent, materialVATType, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function CreateNewDocument(ByVal refDocumentID As Integer, ByVal refDocumentShopID As Integer, ByVal refDocumentStatus As Integer, ByVal inventoryID As Integer,
                                      ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        Dim dtDoc As New DataTable
        Dim dtDocRef As New DataTable
        Dim documentId As Integer = 0
        Dim documentShopId As Integer = 0
        Dim deliveryCost As Decimal = 0

        dtDocRef = DocumentSQL.GetDocument(globalVariable.DocDBUtil, globalVariable.DocConn, refDocumentID, refDocumentShopID, globalVariable.DocLangID)
        If dtDocRef.Rows.Count > 0 Then
            Select Case dtDocRef.Rows(0)("DocumentStatus")
                Case Is = globalVariable.DOCUMENTSTATUS_APPROVE
                    If DocumentModule.CreateNewDocumentFromReferedDocument(globalVariable, globalVariable.DOCUMENTTYPE_ROPO, inventoryID, refDocumentID, refDocumentShopID, Date.Now, True, docData, resultText) = False Then
                        Return False
                    End If

                    If ReceiveFromPurchaseOrderModule.SaveDocumentDataIntoDB(globalVariable, docData.DocumentID, inventoryID, globalVariable.DOCUMENTTYPE_ROPO, Date.Now, docData.VendorID, docData.VendorGroupID, docData.DocumentNote, docData.DocumentRefNumber, deliveryCost, resultText) = False Then
                        Return False
                    End If
                    documentId = docData.DocumentID
                    documentShopId = docData.DocumentShopID
                Case Is = globalVariable.DOCUMENTSTATUS_REFERED
                    dtDoc = DocumentSQL.GetDocumentIDRef(globalVariable.DocDBUtil, globalVariable.DocConn, globalVariable.DOCUMENTTYPE_ROPO, refDocumentID, refDocumentShopID)
                    If dtDoc.Rows.Count > 0 Then
                        documentId = dtDoc.Rows(0)("DocumentId")
                        documentShopId = dtDoc.Rows(0)("ShopId")
                    Else
                        If DocumentModule.CreateNewDocumentFromReferedDocument(globalVariable, globalVariable.DOCUMENTTYPE_ROPO, inventoryID, refDocumentID, refDocumentShopID, Date.Now, True, docData, resultText) = False Then
                            Return False
                        End If
                        If ReceiveFromPurchaseOrderModule.SaveDocumentDataIntoDB(globalVariable, docData.DocumentID, inventoryID, globalVariable.DOCUMENTTYPE_ROPO, Date.Now, docData.VendorID, docData.VendorGroupID, docData.DocumentNote, docData.DocumentRefNumber, deliveryCost, resultText) = False Then
                            Return False
                        End If
                    End If
            End Select
        End If
        
        Return DocumentModule.LoadDocument(globalVariable, documentId, documentShopId, docData, resultText)
    End Function

    Public Function CancelDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromPurchaseOrderModule.CancelDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function DeleteMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal strDocDetailId As String,
                                              ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromPurchaseOrderModule.DeleteDocDetail(globalVariable, documentID, documentShopID, strDocDetailId, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function FinishPurchaseOrderDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef resultText As String) As Boolean
        Dim dbTrans As SqlTransaction
        Dim strUpdateDate As String = FormatDateTime(Date.Now)
        dbTrans = globalVariable.DocConn.BeginTransaction
        Try
            DocumentSQL.FinishDocument(globalVariable.DocDBUtil, globalVariable.DocConn, dbTrans, documentID, documentShopID, strUpdateDate)
            dbTrans.Commit()
        Catch ex As Exception
            dbTrans.Rollback()
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function SaveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date,
                                 ByVal vendorID As Integer, ByVal vendorGroupID As Integer, ByVal documentNote As String, ByVal invoiceReference As String,
                                 ByVal deliveryCost As Decimal, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromPurchaseOrderModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, globalVariable.DOCUMENTTYPE_ROPO, documentDate, vendorID,
                                                                 vendorGroupID, documentNote, invoiceReference, deliveryCost, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, inventoryID, docData, resultText)
    End Function

    Public Function SearchDocument(ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date, ByVal inventoryID As Integer, ByVal vendorID As Integer,
                                   ByVal vendorGroupID As Integer, ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean
        Return ReceiveFromPurchaseOrderModule.SearchDocument(globalVariable, documentStatus, startDate, endDate, inventoryID, vendorID, vendorGroupID, docList, resultText)
    End Function

    Public Function SearchPurchaseOrderDocument(ByVal startDate As Date, ByVal endDate As Date, ByVal inventoryID As Integer,
                                                ByRef docList As List(Of SearchDocumentResult_Data),
                                                ByRef resultText As String) As Boolean
        Return ReceiveFromPurchaseOrderModule.SearchPurchaseOrderDocument(globalVariable, startDate, endDate, inventoryID, docList, resultText)
    End Function

    Public Function UpdateMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer, ByVal materialID As Integer,
                                              ByVal materailQty As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal, ByVal discountAmount As Decimal,
                                              ByVal discountPercent As Decimal, ByVal materialVATType As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If ReceiveFromPurchaseOrderModule.UpdateDocDetail(globalVariable, documentID, documentShopID, docDetailId, materialID, materailQty, materialUnitLargeID,
                                                          pricePerUnit, discountAmount, discountPercent, materialVATType, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

End Class
