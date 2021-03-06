﻿Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient
Imports pRoMiSe.Utilitys.Utilitys

Public Class MonthlyStockController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function AutoTransferStock(ByVal documentDate As Date, ByVal inventoryId As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        Dim documentId As Integer = 0
        Dim endMonth As Date
        Dim dtDocMonth As New DataTable
      
        dtDocMonth = CheckMonthlyDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentDate.Month, documentDate.Year, inventoryId)
        If dtDocMonth.Rows.Count > 0 Then
            documentID = dtDocMonth.Rows(0)("DocumentId")
            inventoryId = dtDocMonth.Rows(0)("ShopId")
            Return DocumentModule.LoadDocument(globalVariable, documentID, inventoryId, docData, resultText)
        End If
        endMonth = New Date(documentDate.Year, documentDate.Month, DateTime.DaysInMonth(documentDate.Year, documentDate.Month))
        If documentDate <> endMonth Then
            resultText = globalVariable.MESSAGE_INVALIDDATECOUNTSTOCK
            Return False
        End If
        If CountStockModule.AutoTransferStock(globalVariable, documentDate, inventoryId, documentId, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentId, inventoryId, docData, resultText)
    End Function

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        Dim stockNotEnoughStock As New List(Of MaterialNotEnoughStock_Data)
        'PTT ไม่ต้องการให้ตรวจสอบสินค้าสต๊อกติดลบ
        'If CheckMaterialStockBelowZero(documentID, documentShopID, Date.Now, stockNotEnoughStock, resultText) = False Then
        '    DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
        '    docData.MaterialNotEnoughStock = stockNotEnoughStock
        '    resultText = globalVariable.MESSAGE_MATERIALBELOWZERO
        '    Return False
        'End If
        If CountStockModule.ApproveDocument(globalVariable, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function ApproveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer,
                                    ByVal documentDate As Date, ByVal documentNote As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean

        Dim stockNotEnoughStock As New List(Of MaterialNotEnoughStock_Data)
        'PTT ไม่ต้องการให้ตรวจสอบสินค้าสต๊อกติดลบ
        'If CheckMaterialStockBelowZero(documentID, documentShopID, documentDate, stockNotEnoughStock, resultText) = False Then
        '    DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
        '    docData.MaterialNotEnoughStock = stockNotEnoughStock
        '    resultText = globalVariable.MESSAGE_MATERIALBELOWZERO
        '    Return False
        'End If
        If CountStockModule.SaveDocumentDataIntoDB(globalVariable, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, documentID, documentShopID, inventoryID, documentDate, documentNote, resultText) = False Then
            Return False
        End If
        If CountStockModule.ApproveDocument(globalVariable, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function AddMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal materialID As Integer, ByVal addAmount As Decimal,
                                           ByVal materialUnitLargeID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If CountStockModule.AddDocDetail(globalVariable, documentID, documentShopID, materialID, addAmount, materialUnitLargeID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function CreateNewDocument(ByVal inventoryID As Integer, ByVal documentDate As Date, ByVal documentNote As String,
                                      ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        Dim endMonth As Date
        Dim dtDocMonth As New DataTable
        Dim documentID, documentShopId As Integer

        dtDocMonth = CheckMonthlyDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentDate.Month, documentDate.Year, inventoryID)
        If dtDocMonth.Rows.Count > 0 Then
            documentID = dtDocMonth.Rows(0)("DocumentId")
            documentShopId = dtDocMonth.Rows(0)("ShopId")
            Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopId, docData, resultText)
        End If
        endMonth = New Date(documentDate.Year, documentDate.Month, DateTime.DaysInMonth(documentDate.Year, documentDate.Month))
        If documentDate <> endMonth Then
            resultText = globalVariable.MESSAGE_INVALIDDATECOUNTSTOCK
            Return False
        End If
        If DocumentModule.CreateNewDocument(globalVariable, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, inventoryID, inventoryID, documentDate, docData, resultText) = False Then
            Return False
        End If
        If CountStockModule.SaveDocumentDataIntoDB(globalVariable, docData.DocumentID, docData.DocumentShopID, inventoryID, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, documentDate,
                                                     documentNote, resultText) = False Then
            Return False
        End If
        If CountStockModule.AutoAddDocDetail(globalVariable, docData.DocumentID, docData.DocumentShopID, "MonthlyStockMaterial", resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, docData.DocumentID, docData.DocumentShopID, docData, resultText)
    End Function

    Public Function CheckMaterialStockBelowZero(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal checkStockDate As Date,
                                              ByRef notEnoughStockDetail As List(Of MaterialNotEnoughStock_Data), ByRef resultText As String) As Boolean

        Dim dtNotEnoughStock As New DataTable
        Dim isFromMatchingTable As Boolean = False
        Dim listOnlyEnoughStock As Boolean = True
        Dim dtResult As New DataTable
        Dim startDate As Date
        Dim endDate As Date
        startDate = New Date(checkStockDate.Year, checkStockDate.Month, 1)
        endDate = checkStockDate
        notEnoughStockDetail = New List(Of MaterialNotEnoughStock_Data)

        If CountStockModule.CheckMaterialStockBelowZero(globalVariable, documentID, documentShopID, startDate, endDate, notEnoughStockDetail, resultText) = True Then
            Return False
        Else
            resultText = ""
            Return True
        End If
    End Function

    Public Function CancelDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If CountStockModule.CancelDocument(globalVariable, documentID, documentShopID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function DeleteMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal strDocDetailId As String, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If CountStockModule.DeleteDocDetail(globalVariable, documentID, documentShopID, strDocDetailId, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

    Public Function SearchDocument(ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date,
                                   ByVal inventoryId As Integer, ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean
        Return CountStockModule.SearchDocument(globalVariable, documentStatus, startDate, endDate, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, inventoryId, docList, resultText)
    End Function

    Public Function SaveDocument(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal inventoryID As Integer,
                                 ByVal documentDate As Date, ByVal documentNote As String, ByRef docData As Document_Data,
                                 ByRef resultText As String) As Boolean
        If CountStockModule.SaveDocumentDataIntoDB(globalVariable, globalVariable.DOCUMENTTYPE_MONTHLYSTOCK, documentID, documentShopID, inventoryID, documentDate, documentNote, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, inventoryID, docData, resultText)
    End Function

    Public Function UpdateMaterialInDocDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal docDetailId As Integer, ByVal materialID As Integer,
                                              ByVal addAmount As Decimal, ByVal materialUnitLargeID As Integer, ByRef docData As Document_Data, ByRef resultText As String) As Boolean
        If CountStockModule.UpdateDocDetail(globalVariable, documentID, documentShopID, docDetailId, materialID, addAmount, materialUnitLargeID, resultText) = False Then
            Return False
        End If
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, docData, resultText)
    End Function

End Class
