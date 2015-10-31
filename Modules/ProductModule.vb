Imports System.Data.SqlClient

Module ProductModule

    Friend Function ProductUpdateData(ByRef ResponseText As String, ByVal ShopID As Integer, ByVal InventoryID As Integer, ByVal ProductGroupID As Integer, ByVal ProductDeptID As Integer, ByVal ProductID As Integer, ByVal MaterialCode As String, ByVal ProductCode As String, ByVal MaterialName As String, ByVal ProductName As String, ByVal MaterialTypeID As Integer, ByVal ProductTypeID As Integer, ByVal MaterialTaxType As Integer, ByVal ProductTaxType As Integer, ByVal ProductDisplay As Integer, ByVal ProductActivate As Integer, ByVal ProductOrdering As Integer, ByVal ProductMin As Double, ByVal ProductPrice0 As Double, ByVal ProductPrice1 As Double, ByVal ProductPrice2 As Double, ByVal ProductPrice3 As Double, ByVal ProductPrice4 As Double, ByVal ProductPrice5 As Double, ByVal ProductPrice6 As Double, ByVal ProductPrice7 As Double, ByVal ProductPrice8 As Double, ByVal ProductPrice9 As Double, ByVal ProductPrice10 As Double, ByVal ProductPrice11 As Double, ByVal ProductPrice12 As Double, ByVal ProductPrice13 As Double, ByVal UnitSmallName As String, ByVal UnitLargeName1 As String, ByVal UnitLargeName2 As String, ByVal UnitLargeName3 As String, ByVal UnitLargeName4 As String, ByVal UnitLargeName5 As String, ByVal UnitRatio1 As Double, ByVal UnitRatio2 As Double, ByVal UnitRatio3 As Double, ByVal UnitRatio4 As Double, ByVal UnitRatio5 As Double, ByVal UnitSmallBarcode As String, ByVal UnitLBarcode1 As String, ByVal UnitLBarcode2 As String, ByVal UnitLBarcode3 As String, ByVal UnitLBarcode4 As String, ByVal UnitLBarcode5 As String, ByVal PurchaseUnit As String, ByVal PurchaseUnit_D As String, ByVal ReceiveUnit As String, ByVal ReceiveUnit_D As String, ByVal SaleUnit As String, ByVal SaleUnit_D As String, ByVal TransferUnit As String, ByVal TransferUnit_D As String, ByVal AdjustUnit As String, ByVal AdjustUnit_D As String, ByVal StockUnit As String, ByVal StockUnit_D As String, ByVal MinimumStock As Decimal, ByVal MaximumStock As Decimal, ByVal IsShowInPOS As Integer, ByVal IsRecommend As Integer, ByVal globalVariable As GlobalVariable) As Boolean
        'Validate Data Section
        Dim FoundError As Boolean = False
        Dim EMsg As String = ""
        If Trim(ProductCode) = "" Then
            FoundError = True
            EMsg = "Product Code must be input before submission"
        End If
        If Trim(MaterialName) = "" Then
            FoundError = True
            EMsg = "Product Name must be input before submission"
        End If
        If Trim(ProductName) = "" Then
            FoundError = True
            EMsg = "Product Alias Name must be input before submission"
        End If
        If Not IsNumeric(ProductPrice1) Then
            FoundError = True
            EMsg = "Product Price 1 must be input and numerice number"
        End If
        If Trim(UnitSmallName) = "" Then
            FoundError = True
            EMsg = "Unit Small Name must be input before submission"
        End If
        Dim objTrx As SqlTransaction
        Dim KeyProductID As Integer = ProductID
        Dim KeyMaterialID As Integer
        Dim KeyUnitSmallID As Integer
        Dim KeyUnitLargeID As Integer
        Dim KeyUnitRatioID As Integer
        Dim KeyPriceID As Integer
        Dim KeyPGroupID As Integer
        Dim productsql, materialsql, unitsmallsql, unitratiosql, unitlargesql, pricesql, componentgroupsql, componentsql As String
        Dim unitlargesql1, unitlargesql2, unitlargesql3, unitlargesql4, unitlargesql5 As String
        Dim unitratiosql1, unitratiosql2, unitratiosql3, unitratiosql4, unitratiosql5 As String
        Dim pricesql1, pricesql2, pricesql3, pricesql4, pricesql5, pricesql6, pricesql7, pricesql8, pricesql9, pricesql10, pricesql11, pricesql12, pricesql13 As String
        Dim SetUnit() As String
        Dim i As Integer
        Dim SettingGroupID As Integer
        Dim SelUnitL(6) As String
        Dim IsDefault As Integer
        Dim IndexUnit As Integer
        Dim TestString As String

        Dim strStep As String = ""
        If FoundError = True Then
            ResponseText = EMsg
            Return False
        Else
            objTrx = globalVariable.DocConn.BeginTransaction(IsolationLevel.Serializable)
            Try
                'Case Add material and product to database
                If ProductID = 0 Then
                    GetMaxID(ResponseText, KeyProductID, 1, globalVariable, objTrx) 'get max productid from table products
                    productsql = "insert into Products (ShopID,InventoryID,ProductGroupID,ProductDeptID,ProductID,ProductCode,ProductName,ProductBarcode,ProductUnitName,ProductTypeID,VATType,ProductDisplay,ProductActivate,ProductOrdering,Deleted,MinimumStock,MaximumForRefillStock,IsShowInPOS,IsRecommend) values (" + ShopID.ToString + "," + InventoryID.ToString + "," + ProductGroupID.ToString + "," + ProductDeptID.ToString + "," + KeyProductID.ToString + ",'" + Replace(ProductCode, "'", "''") + "','" + Replace(ProductName, "'", "''") + "','" + Replace(UnitSmallBarcode, "'", "''") + "','" + Replace(UnitSmallName, "'", "''") + "'," + ProductTypeID.ToString + "," + ProductTaxType.ToString + "," + ProductDisplay.ToString + "," + ProductActivate.ToString + "," + ProductOrdering.ToString + ",0," + MinimumStock.ToString + "," + MaximumStock.ToString + "," + IsShowInPOS.ToString + "," + IsRecommend.ToString + ")"
                    globalVariable.DocDBUtil.sqlExecute(productsql, globalVariable.DocConn, objTrx)
                    TestString += "<br>" + productsql

                    GetMaxID(ResponseText, KeyMaterialID, 2, globalVariable, objTrx) 'get max materialid from table materials
                    GetMaxID(ResponseText, KeyUnitSmallID, 3, globalVariable, objTrx) 'get max unitsmallid from table unitsmall
                    GetMaxID(ResponseText, KeyUnitLargeID, 4, globalVariable, objTrx) 'get max unitlargeid from table unitlarge
                    GetMaxID(ResponseText, KeyUnitRatioID, 5, globalVariable, objTrx) 'get max unitid from table unitratio
                    GetMaxID(ResponseText, KeyPriceID, 6, globalVariable, objTrx) 'get max priceid from table productprice
                    GetMaxID(ResponseText, KeyPGroupID, 7, globalVariable, objTrx) 'get max pgroupid from table productcomponetgroup

                    If ProductTypeID = 0 Then 'Auto add component when product is 1:1
                        materialsql = "insert into materials (MaterialID,MaterialDeptID,MaterialCode,MaterialBarcode,MaterialName,MaterialName1,MaterialTypeID,MaterialTaxType,UnitSmallID,Deleted,MinimumStock,MaximumForRefillStock,IsShowInPOS,IsRecommend) values (" + KeyMaterialID.ToString + "," + ProductDeptID.ToString + ",'" + Replace(MaterialCode, "'", "''") + "','" + Replace(UnitSmallBarcode, "'", "''") + "','" + Replace(MaterialName, "'", "''") + "','" + Replace(ProductName, "'", "''") + "'," + MaterialTypeID.ToString + "," + MaterialTaxType.ToString + "," + KeyUnitSmallID.ToString + ",0," + MinimumStock.ToString + "," + MaximumStock.ToString + "," + IsShowInPOS.ToString + "," + IsRecommend.ToString + ")"
                        globalVariable.DocDBUtil.sqlExecute(materialsql, globalVariable.DocConn, objTrx)

                        If Trim(UnitSmallName) <> "" Then
                            unitsmallsql = "insert into unitsmall (UnitSmallID,UnitSmallName) values (" + KeyUnitSmallID.ToString + ",'" + Replace(UnitSmallName, "'", "''") + "')"
                            unitlargesql = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitSmallName, "'", "''") + "')"
                            unitratiosql = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + KeyUnitSmallID.ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + "1" + ",'" + Replace(UnitSmallBarcode, "'", "''") + "',0)"
                            SelUnitL(0) = KeyUnitLargeID
                            KeyUnitRatioID += 1
                            KeyUnitLargeID += 1
                            globalVariable.DocDBUtil.sqlExecute(unitsmallsql, globalVariable.DocConn, objTrx)
                            globalVariable.DocDBUtil.sqlExecute(unitlargesql, globalVariable.DocConn, objTrx)
                            globalVariable.DocDBUtil.sqlExecute(unitratiosql, globalVariable.DocConn, objTrx)
                        End If
                        If Trim(UnitLargeName1) <> "" And UnitRatio1 > 0 Then
                            unitlargesql1 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName1, "'", "''") + "')"
                            unitratiosql1 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + KeyUnitSmallID.ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio1.ToString + ",'" + Replace(UnitLBarcode1, "'", "''") + "',0)"
                            SelUnitL(1) = KeyUnitLargeID
                            KeyUnitRatioID += 1
                            KeyUnitLargeID += 1
                            globalVariable.DocDBUtil.sqlExecute(unitlargesql1, globalVariable.DocConn, objTrx)
                            globalVariable.DocDBUtil.sqlExecute(unitratiosql1, globalVariable.DocConn, objTrx)
                        End If
                        If Trim(UnitLargeName2) <> "" And UnitRatio2 > 0 Then
                            unitlargesql2 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName2, "'", "''") + "')"
                            unitratiosql2 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + KeyUnitSmallID.ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio2.ToString + ",'" + Replace(UnitLBarcode2, "'", "''") + "',0)"
                            SelUnitL(2) = KeyUnitLargeID
                            KeyUnitRatioID += 1
                            KeyUnitLargeID += 1
                            globalVariable.DocDBUtil.sqlExecute(unitlargesql2, globalVariable.DocConn, objTrx)
                            globalVariable.DocDBUtil.sqlExecute(unitratiosql2, globalVariable.DocConn, objTrx)
                        End If
                        If Trim(UnitLargeName3) <> "" And UnitRatio3 > 0 Then
                            unitlargesql3 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName3, "'", "''") + "')"
                            unitratiosql3 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + KeyUnitSmallID.ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio3.ToString + ",'" + Replace(UnitLBarcode3, "'", "''") + "',0)"
                            SelUnitL(3) = KeyUnitLargeID
                            KeyUnitRatioID += 1
                            KeyUnitLargeID += 1
                            globalVariable.DocDBUtil.sqlExecute(unitlargesql3, globalVariable.DocConn, objTrx)
                            globalVariable.DocDBUtil.sqlExecute(unitratiosql3, globalVariable.DocConn, objTrx)
                        End If
                        If Trim(UnitLargeName4) <> "" And UnitRatio4 > 0 Then
                            unitlargesql4 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName4, "'", "''") + "')"
                            unitratiosql4 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + KeyUnitSmallID.ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio4.ToString + ",'" + Replace(UnitLBarcode4, "'", "''") + "',0)"
                            SelUnitL(4) = KeyUnitLargeID
                            KeyUnitRatioID += 1
                            KeyUnitLargeID += 1
                            globalVariable.DocDBUtil.sqlExecute(unitlargesql4, globalVariable.DocConn, objTrx)
                            globalVariable.DocDBUtil.sqlExecute(unitratiosql4, globalVariable.DocConn, objTrx)
                        End If
                        If Trim(UnitLargeName5) <> "" And UnitRatio1 > 0 Then
                            unitlargesql5 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName5, "'", "''") + "')"
                            unitratiosql5 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + KeyUnitSmallID.ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio5.ToString + ",'" + Replace(UnitLBarcode5, "'", "''") + "',0)"
                            SelUnitL(5) = KeyUnitLargeID
                            KeyUnitRatioID += 1
                            KeyUnitLargeID += 1
                            globalVariable.DocDBUtil.sqlExecute(unitlargesql5, globalVariable.DocConn, objTrx)
                            globalVariable.DocDBUtil.sqlExecute(unitratiosql5, globalVariable.DocConn, objTrx)
                        End If

                        componentgroupsql = "insert into productcomponentgroup (PGroupID,ProductID,SaleMode,SetGroupName) values (" + KeyPGroupID.ToString + "," + KeyProductID.ToString + "," + "1" + ",'Auto add 1:1')"
                        componentsql = "insert into productcomponent (PGroupID,ProductID,SaleMode,MaterialID,MaterialAmount,UnitSmallID) values (" + KeyPGroupID.ToString + "," + KeyProductID.ToString + "," + "1" + "," + KeyMaterialID.ToString + "," + "1" + "," + KeyUnitSmallID.ToString + ")"
                        globalVariable.DocDBUtil.sqlExecute(componentgroupsql, globalVariable.DocConn, objTrx)
                        globalVariable.DocDBUtil.sqlExecute(componentsql, globalVariable.DocConn, objTrx)
                    End If

                    If ProductPrice0 >= 0 Then
                        pricesql = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "0" + "," + ProductPrice0.ToString + ",1)"
                    Else
                        pricesql = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "0" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)

                    If ProductPrice1 >= 0 Then
                        pricesql1 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "1" + "," + ProductPrice1.ToString + ",1)"
                    Else
                        pricesql1 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "1" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql1, globalVariable.DocConn, objTrx)

                    If ProductPrice2 >= 0 Then
                        pricesql2 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "2" + "," + ProductPrice2.ToString + ",1)"
                    Else
                        pricesql2 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "2" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql2, globalVariable.DocConn, objTrx)

                    If ProductPrice3 >= 0 Then
                        pricesql3 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "3" + "," + ProductPrice3.ToString + ",1)"
                    Else
                        pricesql3 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "3" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql3, globalVariable.DocConn, objTrx)

                    If ProductPrice4 >= 0 Then
                        pricesql4 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "4" + "," + ProductPrice4.ToString + ",1)"
                    Else
                        pricesql4 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "4" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql4, globalVariable.DocConn, objTrx)

                    If ProductPrice5 >= 0 Then
                        pricesql5 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "5" + "," + ProductPrice5.ToString + ",1)"
                    Else
                        pricesql5 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "5" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql5, globalVariable.DocConn, objTrx)

                    If ProductPrice6 >= 0 Then
                        pricesql6 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "6" + "," + ProductPrice6.ToString + ",1)"
                    Else
                        pricesql6 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "6" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql6, globalVariable.DocConn, objTrx)

                    If ProductPrice7 >= 0 Then
                        pricesql7 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "7" + "," + ProductPrice7.ToString + ",1)"
                    Else
                        pricesql7 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "7" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql7, globalVariable.DocConn, objTrx)

                    If ProductPrice8 >= 0 Then
                        pricesql8 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "8" + "," + ProductPrice8.ToString + ",1)"
                    Else
                        pricesql8 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "8" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql8, globalVariable.DocConn, objTrx)

                    If ProductPrice9 >= 0 Then
                        pricesql9 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "9" + "," + ProductPrice9.ToString + ",1)"
                    Else
                        pricesql9 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "9" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql9, globalVariable.DocConn, objTrx)

                    If ProductPrice10 >= 0 Then
                        pricesql10 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "10" + "," + ProductPrice10.ToString + ",1)"
                    Else
                        pricesql10 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "10" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql10, globalVariable.DocConn, objTrx)

                    If ProductPrice11 >= 0 Then
                        pricesql11 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "11" + "," + ProductPrice11.ToString + ",1)"
                    Else
                        pricesql11 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "11" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql11, globalVariable.DocConn, objTrx)

                    If ProductPrice12 >= 0 Then
                        pricesql12 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "12" + "," + ProductPrice12.ToString + ",1)"
                    Else
                        pricesql12 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "12" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql12, globalVariable.DocConn, objTrx)

                    If ProductPrice13 >= 0 Then
                        pricesql13 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "13" + "," + ProductPrice13.ToString + ",1)"
                    Else
                        pricesql13 = "insert into ProductPrice (ProductPriceID,ProductID,MainPrice,ProductPrice,SaleMode) values (" + KeyPriceID.ToString + "," + KeyProductID.ToString + "," + "13" + "," + "NULL" + ",1)"
                    End If
                    KeyPriceID += 1
                    globalVariable.DocDBUtil.sqlExecute(pricesql13, globalVariable.DocConn, objTrx)

                    SettingGroupID = 1
                    SetUnit = PurchaseUnit.Split(","c)
                    globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=1 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                    If SetUnit.Length > 0 Then
                        For i = 0 To SetUnit.Length - 1
                            IndexUnit = CInt(SetUnit(i))
                            If SetUnit(i) = PurchaseUnit_D Then
                                IsDefault = 1
                            Else
                                IsDefault = 0
                            End If
                            globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "1" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                            SettingGroupID += 1
                        Next
                    End If

                    SettingGroupID = 1
                    SetUnit = PurchaseUnit.Split(","c)
                    globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=2 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                    If SetUnit.Length > 0 Then
                        For i = 0 To SetUnit.Length - 1
                            IndexUnit = CInt(SetUnit(i))
                            If SetUnit(i) = PurchaseUnit_D Then
                                IsDefault = 1
                            Else
                                IsDefault = 0
                            End If
                            globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "2" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                            SettingGroupID += 1
                        Next
                    End If

                    SettingGroupID = 1
                    SetUnit = ReceiveUnit.Split(","c)
                    globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=39 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                    If SetUnit.Length > 0 Then
                        For i = 0 To SetUnit.Length - 1
                            IndexUnit = CInt(SetUnit(i))
                            If SetUnit(i) = ReceiveUnit_D Then
                                IsDefault = 1
                            Else
                                IsDefault = 0
                            End If
                            globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "39" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                            SettingGroupID += 1
                        Next
                    End If

                    SettingGroupID = 1
                    SetUnit = TransferUnit.Split(","c)
                    globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=3 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                    If SetUnit.Length > 0 Then
                        For i = 0 To SetUnit.Length - 1
                            IndexUnit = CInt(SetUnit(i))
                            If SetUnit(i) = TransferUnit_D Then
                                IsDefault = 1
                            Else
                                IsDefault = 0
                            End If
                            globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "3" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                            SettingGroupID += 1
                        Next
                    End If

                    SettingGroupID = 1
                    SetUnit = AdjustUnit.Split(","c)
                    globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=0 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                    If SetUnit.Length > 0 Then
                        For i = 0 To SetUnit.Length - 1
                            IndexUnit = CInt(SetUnit(i))
                            If SetUnit(i) = AdjustUnit_D Then
                                IsDefault = 1
                            Else
                                IsDefault = 0
                            End If
                            globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "0" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                            SettingGroupID += 1
                        Next
                    End If

                    SettingGroupID = 1
                    SetUnit = StockUnit.Split(","c)
                    globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=7 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                    If SetUnit.Length > 0 Then
                        For i = 0 To SetUnit.Length - 1
                            IndexUnit = CInt(SetUnit(i))
                            If SetUnit(i) = StockUnit_D Then
                                IsDefault = 1
                            Else
                                IsDefault = 0
                            End If
                            globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "7" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                            SettingGroupID += 1
                        Next
                    End If

                Else
                    Dim getInfo As DataTable
                    strStep = "1"
                    getInfo = globalVariable.DocDBUtil.List("select * from products where ProductID=" + ProductID.ToString, globalVariable.DocConn, objTrx)
                    GetMaxID(ResponseText, KeyUnitLargeID, 4, globalVariable, objTrx) 'get max unitlargeid from table unitlarge
                    GetMaxID(ResponseText, KeyUnitRatioID, 5, globalVariable, objTrx) 'get max unitid from table unitratio

                    If getInfo.Rows.Count > 0 Then
                        strStep = "2"
                        productsql = "update products set ProductGroupID=" + ProductGroupID.ToString + ",ProductDeptID=" + ProductDeptID.ToString + ",ProductCode='" + Replace(ProductCode, "'", "''") + "',ProductName='" + Replace(ProductName, "'", "''") + "',ProductBarcode='" + Replace(UnitSmallBarcode, "'", "''") + "',ProductUnitName='" + Replace(UnitSmallName, "'", "''") + "',VATType=" + ProductTaxType.ToString + ",ProductDisplay=" + ProductDisplay.ToString + ",ProductActivate=" + ProductActivate.ToString + ",ProductOrdering=" + ProductOrdering.ToString + ",UpdateDate=getdate(),MinimumStock=" + MinimumStock.ToString + ",MaximumForRefillStock=" + MaximumStock.ToString + ",IsShowInPOS=" + IsShowInPOS.ToString + ",IsRecommend=" + IsRecommend.ToString + " where ProductID=" + ProductID.ToString
                        globalVariable.DocDBUtil.sqlExecute(productsql, globalVariable.DocConn, objTrx)
                        strStep = "3"
                        If ProductPrice0 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice0.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=0"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=0"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice1 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice1.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=1"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=1"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice2 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice2.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=2"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=2"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice3 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice3.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=3"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=3"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice4 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice4.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=4"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=4"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice5 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice5.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=5"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=5"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice6 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice6.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=6"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=6"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice7 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice7.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=7"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=7"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice8 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice8.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=8"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=8"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice9 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice9.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=9"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=9"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice10 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice10.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=10"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=10"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice11 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice11.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=11"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=11"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice12 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice12.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=12"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=12"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If
                        If ProductPrice13 >= 0 Then
                            pricesql = "update ProductPrice set ProductPrice=" + ProductPrice13.ToString + " where ProductID=" + ProductID.ToString + " AND MainPrice=13"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        Else
                            pricesql = "update ProductPrice set ProductPrice=" + "NULL" + " where ProductID=" + ProductID.ToString + " AND MainPrice=13"
                            globalVariable.DocDBUtil.sqlExecute(pricesql, globalVariable.DocConn, objTrx)
                        End If

                        Dim getMat As DataTable
                        Dim unitData As DataTable
                        If getInfo.Rows(0)("ProductTypeID") = 0 Then
                            strStep = "4"
                            getMat = globalVariable.DocDBUtil.List("select * from productcomponent where ProductID=" + ProductID.ToString, globalVariable.DocConn, objTrx)
                            If getMat.Rows.Count > 0 Then
                                KeyUnitSmallID = getMat.Rows(0)("UnitSmallID").ToString
                                KeyMaterialID = getMat.Rows(0)("MaterialID").ToString
                                strStep = "5"
                                materialsql = "update materials set MaterialDeptID=" + ProductDeptID.ToString + ",MaterialCode='" + Replace(MaterialCode, "'", "''") + "',MaterialBarcode='" + Replace(UnitSmallBarcode, "'", "''") + "',MaterialName='" + Replace(MaterialName, "'", "''") + "',MaterialName1='" + Replace(ProductName, "'", "''") + "',MaterialTaxType=" + MaterialTaxType.ToString + ",MinimumStock=" + MinimumStock.ToString + ",MaximumForRefillStock=" + MaximumStock.ToString + ",IsShowInPOS=" + IsShowInPOS.ToString + ",IsRecommend=" + IsRecommend.ToString + " where MaterialID=" + getMat.Rows(0)("MaterialID").ToString
                                globalVariable.DocDBUtil.sqlExecute(materialsql, globalVariable.DocConn, objTrx)
                                strStep = "6"
                                If Trim(UnitSmallName) <> "" Then
                                    unitsmallsql = "update unitsmall set UnitSmallName='" + Replace(UnitSmallName, "'", "''") + "' where UnitSmallID=" + getMat.Rows(0)("UnitSmallID").ToString
                                    globalVariable.DocDBUtil.sqlExecute(unitsmallsql, globalVariable.DocConn, objTrx)
                                End If

                                unitData = globalVariable.DocDBUtil.List("select * from UnitRatio where UnitSmallID=" + getMat.Rows(0)("UnitSmallID").ToString + " order by UnitSmallRatio", globalVariable.DocConn, objTrx)
                                If unitData.Rows.Count > 0 Then
                                    unitlargesql = "update unitlarge set UnitLargeName='" + Replace(UnitSmallName, "'", "''") + "' where UnitLargeID=" + unitData.Rows(0)("UnitLargeID").ToString
                                    unitratiosql = "update UnitRatio set MaterialUnitRatioCode='" + Replace(UnitSmallBarcode, "'", "''") + "' where UnitID=" + unitData.Rows(0)("UnitID").ToString
                                    SelUnitL(0) = unitData.Rows(0)("UnitLargeID").ToString
                                    globalVariable.DocDBUtil.sqlExecute(unitlargesql, globalVariable.DocConn, objTrx)
                                End If
                                strStep = "7"
                                If Trim(UnitLargeName1) <> "" And UnitRatio1 > 0 Then
                                    If unitData.Rows.Count >= 2 Then
                                        unitlargesql1 = "update unitlarge set UnitLargeName='" + Replace(UnitLargeName1, "'", "''") + "' where UnitLargeID=" + unitData.Rows(1)("UnitLargeID").ToString
                                        unitratiosql1 = "update UnitRatio set MaterialUnitRatioCode='" + Replace(UnitLBarcode1, "'", "''") + "',UnitSmallRatio=" + UnitRatio1.ToString + " where UnitID=" + unitData.Rows(1)("UnitID").ToString
                                        SelUnitL(1) = unitData.Rows(1)("UnitLargeID").ToString
                                    Else
                                        unitlargesql1 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName1, "'", "''") + "')"
                                        unitratiosql1 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + getMat.Rows(0)("UnitSmallID").ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio1.ToString + ",'" + Replace(UnitLBarcode1, "'", "''") + "',0)"
                                        SelUnitL(1) = KeyUnitLargeID.ToString
                                        KeyUnitRatioID += 1
                                        KeyUnitLargeID += 1
                                    End If
                                    globalVariable.DocDBUtil.sqlExecute(unitlargesql1, globalVariable.DocConn, objTrx)
                                    globalVariable.DocDBUtil.sqlExecute(unitratiosql1, globalVariable.DocConn, objTrx)
                                End If
                                strStep = "8"
                                If Trim(UnitLargeName2) <> "" And UnitRatio2 > 0 Then
                                    If unitData.Rows.Count >= 3 Then
                                        unitlargesql2 = "update unitlarge set UnitLargeName='" + Replace(UnitLargeName2, "'", "''") + "' where UnitLargeID=" + unitData.Rows(2)("UnitLargeID").ToString
                                        unitratiosql2 = "update UnitRatio set MaterialUnitRatioCode='" + Replace(UnitLBarcode2, "'", "''") + "',UnitSmallRatio=" + UnitRatio2.ToString + " where UnitID=" + unitData.Rows(2)("UnitID").ToString
                                        SelUnitL(2) = unitData.Rows(2)("UnitLargeID").ToString
                                    Else
                                        unitlargesql2 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName2, "'", "''") + "')"
                                        unitratiosql2 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + getMat.Rows(0)("UnitSmallID").ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio2.ToString + ",'" + Replace(UnitLBarcode2, "'", "''") + "',0)"
                                        SelUnitL(2) = KeyUnitLargeID.ToString
                                        KeyUnitRatioID += 1
                                        KeyUnitLargeID += 1
                                    End If
                                    globalVariable.DocDBUtil.sqlExecute(unitlargesql2, globalVariable.DocConn, objTrx)
                                    globalVariable.DocDBUtil.sqlExecute(unitratiosql2, globalVariable.DocConn, objTrx)
                                End If
                                strStep = "9"
                                If Trim(UnitLargeName3) <> "" And UnitRatio3 > 0 Then
                                    If unitData.Rows.Count >= 4 Then
                                        unitlargesql3 = "update unitlarge set UnitLargeName='" + Replace(UnitLargeName3, "'", "''") + "' where UnitLargeID=" + unitData.Rows(3)("UnitLargeID").ToString
                                        unitratiosql3 = "update UnitRatio set MaterialUnitRatioCode='" + Replace(UnitLBarcode1, "'", "''") + "',UnitSmallRatio=" + UnitRatio3.ToString + " where UnitID=" + unitData.Rows(3)("UnitID").ToString
                                        SelUnitL(3) = unitData.Rows(3)("UnitLargeID").ToString
                                    Else
                                        unitlargesql3 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName3, "'", "''") + "')"
                                        unitratiosql3 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + getMat.Rows(0)("UnitSmallID").ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio3.ToString + ",'" + Replace(UnitLBarcode3, "'", "''") + "',0)"
                                        SelUnitL(3) = KeyUnitLargeID.ToString
                                        KeyUnitRatioID += 1
                                        KeyUnitLargeID += 1
                                    End If
                                    globalVariable.DocDBUtil.sqlExecute(unitlargesql3, globalVariable.DocConn, objTrx)
                                    globalVariable.DocDBUtil.sqlExecute(unitratiosql3, globalVariable.DocConn, objTrx)
                                End If
                                strStep = "10"
                                If Trim(UnitLargeName4) <> "" And UnitRatio4 > 0 Then
                                    If unitData.Rows.Count >= 5 Then
                                        unitlargesql4 = "update unitlarge set UnitLargeName='" + Replace(UnitLargeName4, "'", "''") + "' where UnitLargeID=" + unitData.Rows(4)("UnitLargeID").ToString
                                        unitratiosql4 = "update UnitRatio set MaterialUnitRatioCode='" + Replace(UnitLBarcode4, "'", "''") + "',UnitSmallRatio=" + UnitRatio4.ToString + " where UnitID=" + unitData.Rows(4)("UnitID").ToString
                                        SelUnitL(4) = unitData.Rows(4)("UnitLargeID").ToString
                                    Else
                                        unitlargesql4 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName4, "'", "''") + "')"
                                        unitratiosql4 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + getMat.Rows(0)("UnitSmallID").ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio4.ToString + ",'" + Replace(UnitLBarcode4, "'", "''") + "',0)"
                                        SelUnitL(4) = KeyUnitLargeID.ToString
                                        KeyUnitRatioID += 1
                                        KeyUnitLargeID += 1
                                    End If
                                    globalVariable.DocDBUtil.sqlExecute(unitlargesql4, globalVariable.DocConn, objTrx)
                                    globalVariable.DocDBUtil.sqlExecute(unitratiosql4, globalVariable.DocConn, objTrx)
                                End If
                                strStep = "11"
                                If Trim(UnitLargeName5) <> "" And UnitRatio5 > 0 Then
                                    If unitData.Rows.Count >= 6 Then
                                        unitlargesql5 = "update unitlarge set UnitLargeName='" + Replace(UnitLargeName5, "'", "''") + "' where UnitLargeID=" + unitData.Rows(5)("UnitLargeID").ToString
                                        unitratiosql5 = "update UnitRatio set MaterialUnitRatioCode='" + Replace(UnitLBarcode5, "'", "''") + "',UnitSmallRatio=" + UnitRatio5.ToString + " where UnitID=" + unitData.Rows(5)("UnitID").ToString
                                        SelUnitL(5) = unitData.Rows(5)("UnitLargeID").ToString
                                    Else
                                        unitlargesql5 = "insert into unitlarge (UnitLargeID,UnitLargeName) values (" + KeyUnitLargeID.ToString + ",'" + Replace(UnitLargeName5, "'", "''") + "')"
                                        unitratiosql5 = "insert into unitratio (UnitID,UnitSmallID,UnitLargeID,UnitLargeRatio,UnitSmallRatio,MaterialUnitRatioCode,Deleted) values (" + KeyUnitRatioID.ToString + "," + getMat.Rows(0)("UnitSmallID").ToString + "," + KeyUnitLargeID.ToString + "," + "1" + "," + UnitRatio5.ToString + ",'" + Replace(UnitLBarcode5, "'", "''") + "',0)"
                                        SelUnitL(5) = KeyUnitLargeID.ToString
                                        KeyUnitRatioID += 1
                                        KeyUnitLargeID += 1
                                    End If
                                    globalVariable.DocDBUtil.sqlExecute(unitlargesql5, globalVariable.DocConn, objTrx)
                                    globalVariable.DocDBUtil.sqlExecute(unitratiosql5, globalVariable.DocConn, objTrx)
                                End If
                                strStep = "12"
                                SettingGroupID = 1
                                SetUnit = PurchaseUnit.Split(","c)
                                globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=1 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                                If SetUnit.Length > 0 Then
                                    For i = 0 To SetUnit.Length - 1
                                        IndexUnit = CInt(SetUnit(i))
                                        If SetUnit(i) = PurchaseUnit_D Then
                                            IsDefault = 1
                                        Else
                                            IsDefault = 0
                                        End If
                                        globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "1" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                                        SettingGroupID += 1
                                    Next
                                End If
                                strStep = "13"
                                SettingGroupID = 1
                                SetUnit = PurchaseUnit.Split(","c)
                                globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=2 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                                If SetUnit.Length > 0 Then
                                    For i = 0 To SetUnit.Length - 1
                                        IndexUnit = CInt(SetUnit(i))
                                        If SetUnit(i) = PurchaseUnit_D Then
                                            IsDefault = 1
                                        Else
                                            IsDefault = 0
                                        End If
                                        globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "2" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                                        SettingGroupID += 1
                                    Next
                                End If
                                strStep = "14"
                                SettingGroupID = 1
                                SetUnit = ReceiveUnit.Split(","c)
                                globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=39 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                                If SetUnit.Length > 0 Then
                                    For i = 0 To SetUnit.Length - 1
                                        IndexUnit = CInt(SetUnit(i))
                                        If SetUnit(i) = ReceiveUnit_D Then
                                            IsDefault = 1
                                        Else
                                            IsDefault = 0
                                        End If
                                        globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "39" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                                        SettingGroupID += 1
                                    Next
                                End If
                                strStep = "15"
                                SettingGroupID = 1
                                SetUnit = TransferUnit.Split(","c)
                                globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=3 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                                If SetUnit.Length > 0 Then
                                    For i = 0 To SetUnit.Length - 1
                                        IndexUnit = CInt(SetUnit(i))
                                        If SetUnit(i) = TransferUnit_D Then
                                            IsDefault = 1
                                        Else
                                            IsDefault = 0
                                        End If
                                        globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "3" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                                        SettingGroupID += 1
                                    Next
                                End If
                                strStep = "16"
                                SettingGroupID = 1
                                SetUnit = AdjustUnit.Split(","c)
                                globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=0 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                                If SetUnit.Length > 0 Then
                                    For i = 0 To SetUnit.Length - 1
                                        IndexUnit = CInt(SetUnit(i))
                                        If SetUnit(i) = AdjustUnit_D Then
                                            IsDefault = 1
                                        Else
                                            IsDefault = 0
                                        End If
                                        globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "0" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                                        SettingGroupID += 1
                                    Next
                                End If
                                strStep = "17"
                                SettingGroupID = 1
                                SetUnit = StockUnit.Split(","c)
                                globalVariable.DocDBUtil.sqlExecute("delete from MaterialDocumentTypeUnitSetting where DocumentTypeID=7 AND MaterialID=" + KeyMaterialID.ToString + " AND UnitSmallID=" + KeyUnitSmallID.ToString, globalVariable.DocConn, objTrx)
                                If SetUnit.Length > 0 Then
                                    For i = 0 To SetUnit.Length - 1
                                        IndexUnit = CInt(SetUnit(i))
                                        If SetUnit(i) = StockUnit_D Then
                                            IsDefault = 1
                                        Else
                                            IsDefault = 0
                                        End If
                                        globalVariable.DocDBUtil.sqlExecute("insert into MaterialDocumentTypeUnitSetting (SettingGroupID,DocumentTypeID,MaterialID,SelectUnitLargeID,UnitSmallID,IsDefault) values (" + SettingGroupID.ToString + "," + "7" + "," + KeyMaterialID.ToString + "," + SelUnitL(IndexUnit) + "," + KeyUnitSmallID.ToString + "," + IsDefault.ToString + ")", globalVariable.DocConn, objTrx)
                                        SettingGroupID += 1
                                    Next
                                End If

                            End If
                        End If
                    End If
                End If

                objTrx.Commit()
                Return True
            Catch ex As Exception
                objTrx.Rollback()
                ResponseText = "Step : " & strStep & " " & ex.ToString
                Return False
            End Try
        End If
    End Function

    Friend Function GetMaxID(ByRef ResponseText As String, ByRef ReturnMaxID As Integer, ByVal TableID As Integer, ByVal globalVariable As GlobalVariable, ByVal objTrx As SqlTransaction) As Boolean
        'Try
        Dim TableName, ColumnName As String
        Select Case TableID

            Case 1
                TableName = "Products"
                ColumnName = "ProductID"
            Case 2
                TableName = "Materials"
                ColumnName = "MaterialID"
            Case 3
                TableName = "UnitSmall"
                ColumnName = "UnitSmallID"
            Case 4
                TableName = "UnitLarge"
                ColumnName = "UnitLargeID"
            Case 5
                TableName = "UnitRatio"
                ColumnName = "UnitID"
            Case 6
                TableName = "ProductPrice"
                ColumnName = "ProductPriceID"
            Case 7
                TableName = "ProductComponentGroup"
                ColumnName = "PGroupID"
            Case Else
                TableName = "xx"
                ColumnName = "yy"
        End Select

        Dim getMax As DataTable = globalVariable.DocDBUtil.List("select MAX(" + ColumnName + ") As MaxID from " + TableName, globalVariable.DocConn, objTrx)
        Dim MaxID As Integer = 1
        If Not IsDBNull(getMax.Rows(0)("MaxID")) Then
            MaxID = getMax.Rows(0)("MaxID") + 1
        End If
        ResponseText = ""
        ReturnMaxID = MaxID
        Return True
    End Function
End Module
