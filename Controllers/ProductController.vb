﻿Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Public Class ProductController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub
  
    Function ProductUpdateData(ByRef ResponseText As String, ByVal ShopID As Integer, ByVal InventoryID As Integer, ByVal ProductGroupID As Integer, ByVal ProductDeptID As Integer, ByVal ProductID As Integer, ByVal MaterialCode As String, ByVal ProductCode As String, ByVal MaterialName As String, ByVal ProductName As String, ByVal MaterialTypeID As Integer, ByVal ProductTypeID As Integer, ByVal MaterialTaxType As Integer, ByVal ProductTaxType As Integer, ByVal ProductDisplay As Integer, ByVal ProductActivate As Integer, ByVal ProductOrdering As Integer, ByVal ProductMin As Double, ByVal ProductPrice0 As Double, ByVal ProductPrice1 As Double, ByVal ProductPrice2 As Double, ByVal ProductPrice3 As Double, ByVal ProductPrice4 As Double, ByVal ProductPrice5 As Double, ByVal ProductPrice6 As Double, ByVal ProductPrice7 As Double, ByVal ProductPrice8 As Double, ByVal ProductPrice9 As Double, ByVal ProductPrice10 As Double, ByVal ProductPrice11 As Double, ByVal ProductPrice12 As Double, ByVal ProductPrice13 As Double, ByVal UnitSmallName As String, ByVal UnitLargeName1 As String, ByVal UnitLargeName2 As String, ByVal UnitLargeName3 As String, ByVal UnitLargeName4 As String, ByVal UnitLargeName5 As String, ByVal UnitRatio1 As Double, ByVal UnitRatio2 As Double, ByVal UnitRatio3 As Double, ByVal UnitRatio4 As Double, ByVal UnitRatio5 As Double, ByVal UnitSmallBarcode As String, ByVal UnitLBarcode1 As String, ByVal UnitLBarcode2 As String, ByVal UnitLBarcode3 As String, ByVal UnitLBarcode4 As String, ByVal UnitLBarcode5 As String, ByVal PurchaseUnit As String, ByVal PurchaseUnit_D As String, ByVal ReceiveUnit As String, ByVal ReceiveUnit_D As String, ByVal SaleUnit As String, ByVal SaleUnit_D As String, ByVal TransferUnit As String, ByVal TransferUnit_D As String, ByVal AdjustUnit As String, ByVal AdjustUnit_D As String, ByVal StockUnit As String, ByVal StockUnit_D As String, ByVal MinimumStock As Decimal, ByVal MaximumStock As Decimal, ByVal IsShowInPOS As Integer, ByVal IsRecommend As Integer) As Boolean
        Return ProductModule.ProductUpdateData(ResponseText, ShopID, InventoryID, ProductGroupID, ProductDeptID, ProductID, MaterialCode, ProductCode, MaterialName, ProductName, MaterialTypeID, ProductTypeID, MaterialTaxType, ProductTaxType, ProductDisplay, ProductActivate, ProductOrdering, ProductMin, ProductPrice0, ProductPrice1, ProductPrice2, ProductPrice3, ProductPrice4, ProductPrice5, ProductPrice6, ProductPrice7, ProductPrice8, ProductPrice9, ProductPrice10, ProductPrice11, ProductPrice12, ProductPrice13, UnitSmallName, UnitLargeName1, UnitLargeName2, UnitLargeName3, UnitLargeName4, UnitLargeName5, UnitRatio1, UnitRatio2, UnitRatio3, UnitRatio4, UnitRatio4, UnitSmallBarcode, UnitLBarcode1, UnitLargeName2, UnitLBarcode3, UnitLBarcode4, UnitLBarcode5, PurchaseUnit, PurchaseUnit_D, ReceiveUnit, ReceiveUnit_D, SaleUnit, SaleUnit_D, TransferUnit, TransferUnit_D, AdjustUnit, AdjustUnit_D, StockUnit, StockUnit_D, MinimumStock, MaximumStock, IsShowInPOS, IsRecommend, globalVariable)
    End Function

End Class
