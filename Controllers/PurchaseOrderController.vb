Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Public Class PurchaseOrderController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer,  ByRef docData As Document_Data, ByRef resultText As String) As Boolean
       If PurchaseOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date,
                                                  ByVal termOfPayment As Integer, ByVal creditDay As Integer, ByVal dueDate As DateTime, ByVal vendorID As Integer,
                                                  ByVal vendorGroupID As Integer, ByVal documentNote As String, ByVal invoiceReference As String,
                                                  ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If PurchaseOrderModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, globalVariable.DOCUMENTTYPE_PO, documentDate, termOfPayment, creditDay, dueDate, vendorID, vendorGroupID, documentNote, invoiceReference, resultText) = False Then
            Return False
        End If
        If PurchaseOrderModule.ApproveDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AddMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByVal materialID As Integer,
                                                         ByVal materailQty As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal, ByVal discountAmount As Decimal,
                                                         ByVal discountPercent As Decimal, ByVal materialVATType As Integer, ByRef resultText As String) As Boolean
        If PurchaseOrderModule.AddDocDetail(globalVariable, documentID, documentShopID, materialID, materailQty, materialUnitLargeID, pricePerUnit, discountAmount,
                                            discountPercent, materialVATType, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function CreateNewDocument(ByVal inventoryID As Integer, ByVal documentDate As Date, ByVal termOfPayment As Integer, ByVal creditDay As Integer,
                                                    ByVal dueDate As DateTime, ByVal vendorID As Integer, ByVal vendorGroupID As Integer, ByVal documentNote As String,
                                                    ByVal invoiceReference As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If DocumentModule.CreateNewDocument(globalVariable, globalVariable.DOCUMENTTYPE_PO, inventoryID, inventoryID, documentDate, docData, resultText) = False Then
            Return False
        End If
        If PurchaseOrderModule.SaveDocumentDataIntoDB(globalVariable, docData.DocumentID, inventoryID, globalVariable.DOCUMENTTYPE_PO, documentDate, termOfPayment, creditDay, dueDate, vendorID,
                                                     vendorGroupID, documentNote, invoiceReference, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, docData.DocumentID, docData.DocumentShopID, docData, resultText)
    End Function

    Public Function CancelDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If PurchaseOrderModule.CancelDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function DeleteMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal strDocDetailId As String,
                                                           ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        If PurchaseOrderModule.DeleteDocDetail(globalVariable, documentID, documentShopID, strDocDetailId, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function SearchDocument(ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date, ByVal inventoryID As Integer,
                                                 ByVal vendorID As Integer, ByVal vendorGroupID As Integer, ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean
        Return PurchaseOrderModule.SearchDocument(globalVariable, documentStatus, startDate, endDate, inventoryID, vendorID, vendorGroupID, docList, resultText)
    End Function

    Public Function SaveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer, ByVal documentDate As Date,
                                               ByVal termOfPayment As Integer, ByVal creditDay As Integer, ByVal dueDate As DateTime, ByVal vendorID As Integer,
                                               ByVal vendorGroupID As Integer, ByVal documentNote As String, ByVal invoiceReference As String,
                                               ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If PurchaseOrderModule.SaveDocumentDataIntoDB(globalVariable, documentID, inventoryID, globalVariable.DOCUMENTTYPE_PO, documentDate, termOfPayment, creditDay, dueDate, vendorID, vendorGroupID, documentNote, invoiceReference, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, inventoryID, docData, resultText)
    End Function

    Public Function UpdateMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer, ByVal materialID As Integer,
                                                            ByVal materailQty As Decimal, ByVal materialUnitLargeID As Integer, ByVal pricePerUnit As Decimal, ByVal discountAmount As Decimal,
                                                            ByVal discountPercent As Decimal, ByVal materialVATType As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If PurchaseOrderModule.UpdateDocDetail(globalVariable, documentID, documentShopID, docDetailId, materialID, materailQty, materialUnitLargeID, pricePerUnit, discountAmount, discountPercent, materialVATType, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function
    
End Class
