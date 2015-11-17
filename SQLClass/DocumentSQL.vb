Imports pRoMiSe.DBHelper
Imports pRoMiSe.Utilitys.Utilitys
Imports System.Data.SqlClient
Imports System.Text

Module DocumentSQL

    Public Function AddDocDetailMonthlyStockTransfer(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                                     ByVal documentShopId As Integer, ByVal selMonth As Integer, ByVal selYear As Integer) As Boolean
        Dim strSQL As String
        Dim StartDate, EndDate As String
        Dim StartMonthValue, StartYearValue, EndMonthValue, EndYearValue As Integer
        Dim SelectString, WhereString As String
        Dim i As Integer
        Dim qString As New StringBuilder
        Dim getData As DataTable
        Dim newDate As Date
        Dim CalculateFromShopID As Integer = documentShopId
        Dim rowIndex As Integer = 0
        Dim index As Integer = 0

        If selMonth = 12 Then
            StartMonthValue = selMonth
            EndMonthValue = 1
            StartYearValue = selYear
            EndYearValue = selYear + 1
        Else
            StartMonthValue = selMonth
            EndMonthValue = selMonth + 1
            StartYearValue = selYear
            EndYearValue = selYear
        End If

        newDate = New Date(StartYearValue, StartMonthValue, 1)
        StartDate = FormatDate(newDate)
        newDate = New Date(EndYearValue, EndMonthValue, 1)
        EndDate = FormatDate(newDate)

        ' Get standard cost in case amount is 0 from beginning stock                
        strSQL = "IF OBJECT_ID('DummyMaterialCost_Begin', 'U') IS NOT NULL DROP TABLE DummyMaterialCost_Begin;"
        dbUtil.sqlExecute(strSQL, connection, objTrans)
        dbUtil.sqlExecute("create table DummyMaterialCost_Begin (MaterialID int, BeginningPricePerUnit decimal(18,4), PRIMARY KEY (MaterialID))", connection, objTrans)
        strSQL = "select b.ProductID AS MaterialID,b.ProductPricePerUnit from document a, docdetail b, documenttype c where a.documentid=b.documentid and a.shopid=b.shopid and a.documenttypeid=c.documenttypeid and a.shopid=c.shopid and c.langid=1 and c.calculatestandardprofitloss=1 and a.documenttypeid=10 and a.documentstatus=2 and MONTH(a.documentdate) = " + selMonth.ToString + " AND YEAR(a.documentdate) = " + selYear.ToString + " and a.productlevelid=" + CalculateFromShopID.ToString
        dbUtil.sqlExecute("insert into DummyMaterialCost_Begin " + strSQL, connection, objTrans)

        strSQL = "IF OBJECT_ID('DummyMonthlyStandardCost_Temp', 'U') IS NOT NULL DROP TABLE DummyMonthlyStandardCost_Temp;"
        dbUtil.sqlExecute(strSQL, connection, objTrans)
        dbUtil.sqlExecute("create table DummyMonthlyStandardCost_Temp (MaterialID int, NetPrice decimal(18,4), NetAmount decimal(18,4), PRIMARY KEY  (MaterialID) )", connection, objTrans)
        strSQL = "SELECT b.ProductID As MaterialID,SUM(b.ProductNetPrice*c.MovementInStock) AS NetPrice" + _
             " ,SUM(b.UnitSmallAmount*c.MovementInStock) AS NetAmount " + _
             " FROM Document a, DocDetail b, DocumentType c" + _
             " WHERE a.DocumentID=b.DocumentID " + _
             " AND a.ShopID=b.ShopID " + _
             " AND a.DocumentTypeID=c.DocumentTypeID " + _
             " AND a.ShopID=c.ShopID " + _
             " AND a.DocumentDate>=" + StartDate + _
             " AND a.DocumentDate<" + EndDate + _
             " AND c.LangID=1" + _
             " AND c.CalculateStandardProfitLoss=1 AND a.DocumentStatus=2" + _
             " AND a.ProductLevelID=" + CalculateFromShopID.ToString + _
             " GROUP BY b.ProductID"
        dbUtil.sqlExecute("insert into DummyMonthlyStandardCost_Temp (MaterialID,NetPrice,NetAmount) " + strSQL, connection, objTrans)
        SelectString = ", CASE WHERE e.NetPrice is NULL THEN 0 ELSE e.NetPrice END AS NetPrice, CASE WHERE e.NetAmount is NULL THEN 0 ELSE e.NetAmount END AS NetAmount"
        WhereString = " left outer join DummyMonthlyStandardCost_Temp e ON a.MaterialID=e.MaterialID"

        strSQL = "IF OBJECT_ID('DummyMonthlyStock_Temp', 'U') IS NOT NULL DROP TABLE DummyMonthlyStock_Temp;"
        dbUtil.sqlExecute(strSQL, connection, objTrans)
        dbUtil.sqlExecute("create table DummyMonthlyStock_Temp (MaterialID int, Amount decimal(18,4), PRIMARY KEY  (MaterialID) )", connection, objTrans)
        strSQL = "SELECT b.ProductID As MaterialID" + _
                " ,SUM(b.UnitSmallAmount*c.MovementInStock) AS Amount " + _
             " FROM Document a, DocDetail b, DocumentType c" + _
             " WHERE a.DocumentID=b.DocumentID " + _
             " AND a.ShopID=b.ShopID " + _
             " AND a.DocumentTypeID=c.DocumentTypeID " + _
             " AND a.ShopID=c.ShopID " + _
             " AND a.DocumentDate>=" + StartDate + _
             " AND a.DocumentDate<" + EndDate + _
             " AND c.LangID=1" + _
             " AND a.DocumentStatus=2" + _
             " AND a.ProductLevelID=" + documentShopId.ToString + _
             " GROUP BY b.ProductID"
        dbUtil.sqlExecute("insert into DummyMonthlyStock_Temp (MaterialID,Amount) " + strSQL, connection, objTrans)

        strSQL = "IF OBJECT_ID('DummyDocMonthlyStock_Temp', 'U') IS NOT NULL DROP TABLE DummyDocMonthlyStock_Temp;"
        dbUtil.sqlExecute(strSQL, connection, objTrans)
        dbUtil.sqlExecute("create table DummyDocMonthlyStock_Temp (MaterialID int, RecordAmount decimal(18,4), ProductAmount decimal(18,4), UnitID int, UnitName varchar(100), ProductUnit int, UnitRatio decimal(18,4), PRIMARY KEY  (MaterialID) )", connection, objTrans)
        strSQL = "select b.ProductID As MaterialID, b.UnitSmallAmount As RecordAmount, b.ProductAmount, b.UnitID, b.UnitName, b.ProductUnit, c.UnitSmallRatio from Document a, DocDetail b, UnitRatio c where a.DocumentID=b.DocumentID AND a.ShopID=b.ShopID AND b.ProductUnit=c.UnitSmallID AND b.UnitID=c.UnitLargeID AND a.DocumentTypeID=7 AND a.DocumentStatus=1 AND a.ProductLevelID=" + documentShopId.ToString + " AND a.DocumentMonth=" + selMonth.ToString + " AND a.DocumentYear=" + selYear.ToString
        dbUtil.sqlExecute("insert into DummyDocMonthlyStock_Temp (MaterialID,RecordAmount,ProductAmount,UnitID,UnitName,ProductUnit,UnitRatio) " + strSQL, connection, objTrans)

        strSQL = "IF OBJECT_ID('DummyMonthlyStockTransfer_Temp', 'U') IS NOT NULL DROP TABLE DummyMonthlyStockTransfer_Temp;"
        getData = dbUtil.List(strSQL, connection, objTrans)
        dbUtil.sqlExecute("create table DummyMonthlyStockTransfer_Temp (MaterialID int, UnitSmallID int, UnitSmallName varchar(100), NetPrice decimal(18,4), NetAmount decimal(18,4), BeginningPricePerUnit decimal(18,4), Amount decimal(18,4), RecordAmount decimal(18,4), ProductAmount decimal(18,4), UnitID int, UnitName varchar(100), ProductUnit int, UnitRatio decimal(18,4), PRIMARY KEY  (MaterialID) )", connection, objTrans)
        strSQL = "select a.MaterialID,a.UnitSmallID, u.UnitSmallName, bg.BeginningPricePerUnit," &
                    "CASE WHEN e.NetPrice is NULL THEN 0 ELSE e.NetPrice END AS NetPrice," &
                    "CASE WHEN e.NetAmount is NULL THEN 0 ELSE e.NetAmount END AS NetAmount," &
                    "CASE WHEN c.Amount is NULL THEN 0 ELSE c.Amount END As Amount, " &
                    "CASE WHEN d.RecordAmount is NULL THEN  (CASE WHEN c.Amount < 0 THEN 0 ELSE  (CASE WHEN c.Amount is NULL THEN 0 ELSE c.Amount END) END) ELSE d.RecordAmount END As RecordAmount," &
                    "CASE WHEN d.ProductAmount is NULL THEN ( CASE WHEN c.Amount < 0 THEN 0 ELSE (CASE WHEN c.Amount is NULL THEN 0 ELSE c.Amount END) END) ELSE d.ProductAmount END AS ProductAmount," &
                    "CASE WHEN d.UnitID is NULL THEN a.UnitSmallID ELSE d.UnitID END AS UnitID," &
                    "CASE WHEN d.UnitName is NULL THEN u.UnitSmallName ELSE d.UnitName END AS UnitName," &
                    "CASE WHEN d.ProductUnit is NULL THEN a.UnitSmallID ELSE d.ProductUnit END AS ProductUnit," &
                    "CASE WHEN d.UnitRatio is NULL THEN 1 ELSE d.UnitRatio END AS UnitRatio " &
                    "from Materials a left outer join DummyMonthlyStock_Temp c ON a.MaterialID=c.MaterialID left outer join DummyDocMonthlyStock_Temp d ON a.MaterialID=d.MaterialID " &
                    "left outer join DummyMonthlyStandardCost_Temp e ON a.MaterialID=e.MaterialID left outer join UnitSmall u ON a.UnitSmallID=u.UnitSmallID " &
                    "left outer join DummyMaterialCost_Begin bg ON a.MaterialID=bg.MaterialID where a.Deleted=0 "
        dbUtil.sqlExecute("insert into DummyMonthlyStockTransfer_Temp (MaterialID,UnitSmallID,UnitSmallName,BeginningPricePerUnit,NetPrice,NetAmount,Amount,RecordAmount,ProductAmount,UnitID,UnitName,ProductUnit,UnitRatio) " + strSQL, connection, objTrans)

        strSQL = "select *, CASE WHEN NetAmount = 0 THEN CASE WHEN BeginningPricePerUnit is NULL THEN 0 ELSE BeginningPricePerUnit END ELSE CASE WHEN NetPrice = 0 THEN CASE WHEN BeginningPricePerUnit is NULL THEN 0 ELSE BeginningPricePerUnit END ELSE NetPrice/NetAmount END END AS PricePerUnit from DummyMonthlyStockTransfer_Temp"
        getData = dbUtil.List(strSQL, connection, objTrans)
        qString.Append("INSERT INTO DocDetail (DocDetailID, DocumentID, ShopID, ProductID, ProductUnit, ProductAmount,StockAmount, ProductPricePerUnit, ProductNetPrice, UnitName, UnitSmallAmount, UnitID) VALUES ")

        For i = 0 To getData.Rows.Count - 1
            rowIndex = i + 1
            If rowIndex = 999 Then
                rowIndex = 0
                qString = qString.Append(";")
                qString.Append("INSERT INTO DocDetail (DocDetailID, DocumentID, ShopID, ProductID, ProductUnit, ProductAmount,StockAmount, ProductPricePerUnit, ProductNetPrice, UnitName, UnitSmallAmount, UnitID) VALUES ")
                qString = qString.Append("(" + (i + 1).ToString + "," + documentId.ToString + "," + documentShopId.ToString + "," + getData.Rows(i)("MaterialID").ToString + "," + getData.Rows(i)("ProductUnit").ToString + "," + getData.Rows(i)("ProductAmount").ToString + "," + getData.Rows(i)("ProductAmount").ToString + "," + getData.Rows(i)("PricePerUnit").ToString + "," + (getData.Rows(i)("PricePerUnit") * getData.Rows(i)("RecordAmount")).ToString + ",'" + getData.Rows(i)("UnitName") + "'," + getData.Rows(i)("RecordAmount").ToString + "," + getData.Rows(i)("UnitID").ToString + ")")
            Else
                If i > 0 Then qString = qString.Append(",")
                qString = qString.Append("(" + (i + 1).ToString + "," + documentId.ToString + "," + documentShopId.ToString + "," + getData.Rows(i)("MaterialID").ToString + "," + getData.Rows(i)("ProductUnit").ToString + "," + getData.Rows(i)("ProductAmount").ToString + "," + getData.Rows(i)("ProductAmount").ToString + "," + getData.Rows(i)("PricePerUnit").ToString + "," + (getData.Rows(i)("PricePerUnit") * getData.Rows(i)("RecordAmount")).ToString + ",'" + getData.Rows(i)("UnitName") + "'," + getData.Rows(i)("RecordAmount").ToString + "," + getData.Rows(i)("UnitID").ToString + ")")
            End If
        Next

        strSQL = qString.ToString
        If strSQL <> "" Then
            index = dbUtil.sqlExecute(strSQL, connection, objTrans)
        End If

        Return index
    End Function

    Friend Function ApproveDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                           ByVal documentShopId As Integer, ByVal approveBy As Integer, ByVal updateDate As String) As Integer
        Dim strSQL As String
        strSQL = "Update Document " & _
                 "Set DocumentStatus = 2, ApproveBy = " & approveBy &
                 ", UpdateDate = " & updateDate & ", ApproveDate = " & updateDate &
                 ", AlreadyExportToHQ = 0, AlreadyExportToBranch = 0 " &
                 "Where DocumentID = " & documentId & " AND ShopID = " & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function CreateNewDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction,
                                            ByVal newDocID As Integer, ByVal docShopID As Integer, ByVal productLevelID As Integer,
                                            ByVal toInvID As Integer, ByVal docTypeID As Integer, ByVal docStatus As Integer, ByVal docDate As String,
                                            ByVal dueDate As String, ByVal documentNote As String, ByVal createDate As String, ByVal createBy As Integer) As Integer

        Dim strSQL As String
        strSQL = "Insert INTO Document(DocumentID, ShopID, DocumentTypeID, DocumentStatus, DocumentDate, " &
                 " ProductLevelID, ToInvID,Remark, InputBy, InsertDate) " &
                 " VALUES (" & newDocID & ", " & docShopID & ", " & docTypeID & ", " & docStatus & "," & docDate & ", " &
                 productLevelID & ", " & toInvID & ", '" & ReplaceSuitableStringForSQL(documentNote) & "', " &
                 createBy & ", " & createDate & " ) "
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function CreateNewDocumentFromReferedDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal newDocID As Integer,
                                                         ByVal docShopID As Integer, ByVal documentTypeId As Integer, ByVal copyFromDocumentID As Integer,
                                                         ByVal copyFromDocumentShopID As Integer, ByVal documentNote As String, ByVal createDate As String, ByVal documentDate As String,
                                                         ByVal createBy As Integer) As Integer
        Dim strSQL As String
        strSQL = "Delete From Document Where DocumentID = " & newDocID & " AND ShopID = " & docShopID
        dbUtil.sqlExecute(strSQL, connection, objTrans)

        strSQL = "Insert INTO Document(DocumentID, ShopID, DocumentTypeID, DocumentStatus, DocumentDate, " &
                 " ProductLevelID, FromInvID, DocumentIDRef, DocumentIDRefShopID, DueDate,Remark, invoiceref, InputBy, InsertDate," &
                 "vendorid,vendorgroupid,VATPercent,subtotal,TotalDiscount,TotalVAT,netprice,GrandTotal,TermOfPayment,CreditDay) " &
                 " Select " & newDocID & "," & docShopID & ", " & documentTypeId & ", 1, " & documentDate &
                 ", " & docShopID & "," & copyFromDocumentShopID & ", " & copyFromDocumentID & "," & copyFromDocumentShopID &
                 ", DueDate,Remark,'" & Trim(documentNote) & "'," & createBy & "," & createDate &
                 ",vendorid,vendorgroupid,VATPercent,subtotal,TotalDiscount,TotalVAT,netprice,GrandTotal,TermOfPayment,CreditDay" &
                 " From Document " &
                 " Where DocumentID = " & copyFromDocumentID & " AND ShopID = " & copyFromDocumentShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function CopyDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction,
                                              ByVal copyFromDocumentID As Integer, ByVal copyFromDocumentShopID As Integer, ByVal copyToDocumentID As Integer,
                                              ByVal copyToDocumentShopID As Integer, ByVal docStatus As Integer) As Integer
        Dim strSQL As String
        strSQL = "Delete From DocDetail Where DocumentID = " & copyToDocumentID & " AND ShopID = " & copyToDocumentShopID
        dbUtil.sqlExecute(strSQL, connection, objTrans)

        strSQL = "Insert INTO DocDetail(DocDetailID, DocumentID, ShopID, ProductID, ProductUnit, ProductAmount, ProductCode,ProductName, " &
                 " ProductDiscount, ProductDiscountAmount, ProductPricePerUnit, ProductTax, ProductTaxType, " &
                 " ProductTaxIn, MarkUp, PricePerUnitBeforeMark, UnitName, UnitSmallAmount, UnitID, ProductNetPrice, DefaultInCompare, " &
                 " RequestSmallAmount, PrepareSmallAmount, TransferSmallAmount, ROSmallAmount, ReferenceNetPrice, ReferenceProductTax) "
        If docStatus = 3 Then
            strSQL &= " Select DocDetailID, " & copyToDocumentID & ", " & copyToDocumentShopID & ", ProductID, ProductUnit, ProductAmount, ProductCode, ProductName, " &
               " 0, 0, ProductPricePerUnit, ProductTax, ProductTaxType, " &
               " ProductTaxIn, MarkUp, PricePerUnitBeforeMark, UnitName,UnitSmallAmount, UnitID, ProductNetPrice, 1, " &
               " RequestSmallAmount, PrepareSmallAmount, TransferSmallAmount, ROSmallAmount , ReferenceNetPrice, ReferenceProductTax" &
               " From DocDetail " &
               " Where DocumentID = " & copyFromDocumentID & " AND ShopID = " & copyFromDocumentShopID & " AND UnitSmallAmount <> 0 "
        Else
            strSQL &= " Select DocDetailID, " & copyToDocumentID & ", " & copyToDocumentShopID & ", ProductID, ProductUnit, ProductAmount, ProductCode, ProductName, " &
                 " ProductDiscount, ProductDiscountAmount, ProductPricePerUnit, ProductTax, ProductTaxType, " &
                 " ProductTaxIn, MarkUp, PricePerUnitBeforeMark, UnitName,UnitSmallAmount, UnitID, ProductNetPrice, 1, " &
                 " RequestSmallAmount, PrepareSmallAmount, TransferSmallAmount, ROSmallAmount , ReferenceNetPrice, ReferenceProductTax" &
                 " From DocDetail " &
                 " Where DocumentID = " & copyFromDocumentID & " AND ShopID = " & copyFromDocumentShopID & " AND UnitSmallAmount <> 0 "
        End If

        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function CancelDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal docID As Integer,
                                          ByVal docShopID As Integer, ByVal cancelBy As Integer, ByVal updateDate As String) As Integer
        Dim strSQL As String
        strSQL = "Update Document " & _
                 "Set DocumentStatus = 99, NewSend=1, VoidBy = " & cancelBy & ", UpdateDate = " & updateDate & ", CancelDate = " & updateDate &
                 ", AlreadyExportToHQ = 0, AlreadyExportToBranch = 0 " &
                 "Where DocumentID = " & docID & " AND ShopID = " & docShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function CheckMaterialInStockAndCalculateAveragePricePerUnitForTransfer(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection,
                           ByVal docID As Integer, ByVal docShopID As Integer, ByVal docDate As String, ByVal productLevelID As Integer, ByVal staffID As Integer,
                           ByVal isFromMatchingTable As Boolean, ByVal isAutoCreateDRO As Boolean, ByVal listOnlyNotEnoughStock As Boolean,
                           ByRef dtNotEnoughStock As DataTable) As Boolean

        Dim strSQL, strWhere As String
        Dim dtPrice As DataTable
        Dim strStartDate As String
        Dim i As Integer

        'Find the lastest TransferStock document for this productlevelid
        strSQL = "Select Max(DocumentDate) as DocumentDate " & _
                 "From Document " & _
                 "Where DocumentTypeID = 10 AND DocumentStatus = 2 AND ProductLevelID = " & productLevelID & " AND DocumentDate <= " & docDate
        dtPrice = dbUtil.List(strSQL, connection)

        If dtPrice.Rows.Count = 0 Then
            strStartDate = FormatDate(Now.AddYears(-5).Date)
        ElseIf IsDBNull(dtPrice.Rows(0)("DocumentDate")) Then
            strStartDate = FormatDate(Now.AddYears(-5).Date)
        Else
            strStartDate = FormatDate(dtPrice.Rows(0)("DocumentDate"))
        End If

        '4 Temp Table
        '1 : DocDetailForThisDocument = Table for DocumentDetail for this transfer document
        '2 : SumMaterialThisDocument = Table for summary of material from DocDetailForThisDocument
        '   (Because maybe some material has multiple unit)
        '3 : CurrentMaterailAmountInStock = Table for current material amount (only material in SumMaterialThisDocument).
        '    This will be used to compare with the 2nd table whether there are enough material to transfer and 
        '    it also has its total net price.
        '4 : DocumentForCalculateStock = Table for Document that use in Calculate Stock but it not include document in
        '   SumMaterialThisDocument
        '5 : MaterialForCalculateStock = Table for Material that use in Calculate Stock (Material In SumMaterialThisDocument)
        '********************************************************************************************

        'DocDetailForThisDocument Temp Table
        strSQL = "IF OBJECT_ID('DocDetailForThisDocument" & staffID & "', 'U') IS NOT NULL DROP TABLE DocDetailForThisDocument" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Create Table DocDetailForThisDocument" & staffID & _
                " (DocumentID int NOT NULL, DocShopID int NOT NULL, DocDetailID int NOT NULL, " & _
                "  MaterialID int NOT NULL, MaterialAmount decimal(18,4), UnitSmallAmount decimal(18,4), " & _
                "  UnitID int NOT NULL, UnitSmallID int NOT NULL, ProductNetPrice decimal(18,4)); "
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Insert INTO DocDetailForThisDocument" & staffID & "(DocumentID, DocShopID, DocDetailID, MaterialID, MaterialAmount, UnitSmallAmount, UnitID,UnitSmallID, ProductNetPrice ) " & _
                    " Select DocumentID, ShopID, DocDetailID, ProductID, ProductAmount, UnitSmallAmount, UnitID, " & _
                    " ProductUnit, ProductNetPrice " & _
                    " From DocDetail " & _
                    " Where DocumentID = " & docID & " AND ShopID = " & docShopID
        dbUtil.sqlExecute(strSQL, connection)

        '********************************************************************************************
        'SumMaterialThisDocument Temp Table
        strSQL = "IF OBJECT_ID('SumMaterialThisDocument" & staffID & "', 'U') IS NOT NULL DROP TABLE SumMaterialThisDocument" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Create Table SumMaterialThisDocument" & staffID & _
                " (DocumentID int NOT NULL, DocShopID int NOT NULL, MaterialID int NOT NULL, SummaryAmount decimal(18,4)); "
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Insert INTO SumMaterialThisDocument" & staffID & "(DocumentID, DocShopID, MaterialID, SummaryAmount) " &
               "Select DocumentID, DocShopID, MaterialID,Sum(UnitSmallAmount) " &
               "From DocDetailForThisDocument" & staffID & " " & _
               "Group by DocumentID, DocShopID, MaterialID "
        dbUtil.sqlExecute(strSQL, connection)

        '********************************************************************************************
        'DocumentForCalculateStock Temp Table >> Document that not count DocumentID/ShopID In SumMaterialThisDocument
        strSQL = "IF OBJECT_ID('DocumentForCalculateStock" & staffID & "', 'U') IS NOT NULL DROP TABLE DocumentForCalculateStock" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Create Table DocumentForCalculateStock" & staffID & _
                " (DocumentID int NOT NULL, ShopID int NOT NULL, DocumentTypeID int NOT NULL, " & _
                "  MovementInStock int NOT NULL); "
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Insert INTO DocumentForCalculateStock" & staffID & " " & _
                 "Select d.DocumentID, d.ShopID, d.DocumentTypeID, dt.MovementInStock " & _
                 "From Document d, DocumentType dt, DocumentTypeGroupValue dg " & _
                 "Where d.ProductLevelID = " & productLevelID & " AND DocumentStatus = 2 AND " & _
                 " d.DocumentDate >= " & strStartDate & " AND d.DocumentDate <= " & docDate & " AND dt.LangID = 1 AND dt.ComputerID = 0 AND " & _
                 " d.ShopID = dt.ShopID AND d.DocumentTypeID = dt.DocumentTypeID AND dt.MovementInStock <> 0 AND " & _
                 " dt.DocumentTypeID = dg.DocumentTypeID AND d.DocumentTypeID NOT IN (10) " & _
                 " UNION "
        'Transfer Stock For This Month (if exist)
        strSQL &= "Select d.DocumentID, d.ShopID, d.DocumentTypeID, dt.MovementInStock " & _
                 "From Document d, DocumentType dt " & _
                 "Where d.ProductLevelID = " & productLevelID & " AND DocumentStatus = 2 AND " & _
                 " d.DocumentDate = " & strStartDate & " AND dt.LangID = 1 AND dt.ComputerID = 0 AND " & _
                 " d.ShopID = dt.ShopID AND d.DocumentTypeID = dt.DocumentTypeID AND dt.MovementInStock <> 0 AND " & _
                 " dt.DocumentTypeID = 10 "
        dbUtil.sqlExecute(strSQL, connection)

        strSQL = "Select Distinct DocumentID, DocShopID From SumMaterialThisDocument" & staffID
        dtPrice = dbUtil.List(strSQL, connection)
        'Delete Document from SumMaterialThisDocument in DocumentForCalculateStock --> Only not Auto Create DRO Document
        For i = 0 To dtPrice.Rows.Count - 1
            strSQL = " Delete From DocumentForCalculateStock" & staffID & _
                     " Where DocumentID = " & dtPrice.Rows(i)("DocumentID") & " AND ShopID = " & dtPrice.Rows(i)("DocShopID")
            dbUtil.sqlExecute(strSQL, connection)
        Next i

        '********************************************************************************************
        'MaterialForCalculateStock Temp Table >> Material that are in SumMaterialThisDocument
        strSQL = "IF OBJECT_ID('MaterialForCalculateStock" & staffID & "', 'U') IS NOT NULL DROP TABLE MaterialForCalculateStock" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Create Table MaterialForCalculateStock" & staffID & _
                " (MaterialID int NOT NULL); "
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Insert INTO MaterialForCalculateStock" & staffID & " " & _
                 "Select Distinct MaterialID From SumMaterialThisDocument" & staffID
        dbUtil.sqlExecute(strSQL, connection)

        '********************************************************************************************
        'CurrentAllMaterailInStock Temp Table (Do All Material and then filter it in other table 
        ' (join DocumentForCalculateStock and DocDetail) is faster than
        ' join 3 Table : DocumentForCalculateStock, DocDetail and MaterialForCalculateStock)
        strSQL = "IF OBJECT_ID('CurrentAllMaterailInStock" & staffID & "', 'U') IS NOT NULL DROP TABLE CurrentAllMaterailInStock" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Create Table CurrentAllMaterailInStock" & staffID & _
                " (MaterialID int NOT NULL PRIMARY KEY, CurrentAmount decimal(18,4), MaterialPrice decimal(18,4)); "
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Insert INTO CurrentAllMaterailInStock" & staffID & " " & _
                 "Select dd.ProductID, Sum(d.MovementInStock * dd.UnitSmallAmount), " & _
                 " Sum(d.MovementInStock * dd.ProductNetPrice) " & _
                 "From DocumentForCalculateStock" & staffID & " d, DocDetail dd " & _
                 "Where d.DocumentID = dd.DocumentID AND d.ShopID = dd.ShopID " & _
                 "Group by dd.ProductID "
        dbUtil.sqlExecute(strSQL, connection)

        '********************************************************************************************
        'CurrentMaterailInStock Temp Table
        strSQL = "IF OBJECT_ID('CurrentMaterailInStock" & staffID & "', 'U') IS NOT NULL DROP TABLE CurrentMaterailInStock" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Create Table CurrentMaterailInStock" & staffID & _
                " (MaterialID int NOT NULL, CurrentAmount decimal(18,4), MaterialPrice decimal(18,4)); "
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Insert INTO CurrentMaterailInStock" & staffID & " " & _
                 "Select dd.ProductID, Sum(d.MovementInStock * dd.UnitSmallAmount), " & _
                 " Sum(d.MovementInStock * dd.ProductNetPrice) " & _
                 "From DocumentForCalculateStock" & staffID & " d, DocDetail dd, " & _
                 " MaterialForCalculateStock" & staffID & " m " & _
                 "Where d.DocumentID = dd.DocumentID AND d.ShopID = dd.ShopID AND dd.ProductID = m.MaterialID " & _
                 "Group by dd.ProductID "
        strSQL = "Insert INTO CurrentMaterailInStock" & staffID & " " & _
                 "Select dd.MaterialID, CurrentAmount, MaterialPrice " & _
                 "From CurrentAllMaterailInStock" & staffID & " dd, MaterialForCalculateStock" & staffID & " m " & _
                 "Where dd.MaterialID = m.MaterialID  "
        dbUtil.sqlExecute(strSQL, connection)

        'Find Material that are not enough to transfer
        If isFromMatchingTable = False Then
            'Can Query from SumMaterialThisDocument because there will have only 1 document in table
            If listOnlyNotEnoughStock = True Then
                strWhere = " AND ((cm.CurrentAmount < sm.SummaryAmount) OR (cm.CurrentAmount Is NULL)) AND sm.SummaryAmount <> 0 "
            Else
                strWhere = " "
            End If
            strSQL = "Select sm.MaterialID, m.MaterialCode, m.MaterialName, sm.SummaryAmount as TransferAmount, " & _
                     "  cm.CurrentAmount, us.UnitSmallID, " & _
                     "  us.UnitSmallName, ul.UnitLargeID, ul.UnitLargeName, ur.UnitLargeRatio, ur.UnitSmallRatio " & _
                     "From Materials m JOIN UnitSmall us ON m.UnitSmallID = us.UnitSmallID " & _
                     " JOIN UnitRatio ur ON us.UnitSmallID = ur.UnitSmallID " & _
                     " JOIN UnitLarge ul ON ur.UnitLargeID = ul.UnitLargeID " & _
                     " JOIN SumMaterialThisDocument" & staffID & " sm ON m.MaterialID = sm.MaterialID " & _
                     " LEFT OUTER JOIN CurrentMaterailInStock" & staffID & " cm ON sm.MaterialID = cm.MaterialID " & _
                     "Where ur.Deleted = 0 " & strWhere & _
                     "Order by MaterialName, sm.MaterialID, ur.UnitSmallRatio DESC "
        Else
            'Can not Query from SumMaterialThisDocument because there will have more than 1 document in table
            strSQL = "IF OBJECT_ID('SumMaterialAllDocument" & staffID & "', 'U') IS NOT NULL DROP TABLE SumMaterialAllDocument" & staffID & ";"
            dbUtil.sqlExecute(strSQL, connection)
            'Need to create temp table for summary all data in SumMaterialThisDocument
            strSQL = "Create Table SumMaterialAllDocument" & staffID & _
                    " (MaterialID int NOT NULL, SummaryAmount decimal(18,4)); "
            dbUtil.sqlExecute(strSQL, connection)
            strSQL = "Insert INTO SumMaterialAllDocument" & staffID & " " & _
                     "Select MaterialID, Sum(SummaryAmount) " & _
                     "From SumMaterialThisDocument" & staffID & " Group by MaterialID "
            dbUtil.sqlExecute(strSQL, connection)
            If listOnlyNotEnoughStock = True Then
                strWhere = " AND ((cm.CurrentAmount < sm.SummaryAmount) OR (cm.CurrentAmount Is NULL)) AND sm.SummaryAmount <> 0 "
            Else
                strWhere = " "
            End If
            strSQL = "Select sm.MaterialID, m.MaterialCode, m.MaterialName, sm.SummaryAmount as TransferAmount, " & _
                     "  cm.CurrentAmount, us.UnitSmallID, " & _
                     "  us.UnitSmallName, ul.UnitLargeID, ul.UnitLargeName, ur.UnitLargeRatio, ur.UnitSmallRatio " & _
                     "From Materials m JOIN UnitSmall us ON m.UnitSmallID = us.UnitSmallID " & _
                     " JOIN UnitRatio ur ON us.UnitSmallID = ur.UnitSmallID " & _
                     " JOIN UnitLarge ul ON ur.UnitLargeID = ul.UnitLargeID " & _
                     " JOIN SumMaterialAllDocument" & staffID & " sm ON  m.MaterialID = sm.MaterialID " & _
                     " LEFT OUTER JOIN CurrentMaterailInStock" & staffID & " cm ON sm.MaterialID = cm.MaterialID " & _
                     "Where ur.Deleted = 0 " & strWhere & _
                     "Order by MaterialName, sm.MaterialID, ur.UnitSmallRatio DESC "
        End If
        dtNotEnoughStock = dbUtil.List(strSQL, connection)

        'Some Material are not enough to transfer
        If dtNotEnoughStock.Rows.Count > 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Friend Function CrateDummyMaterialStdCost(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal ShopID As Integer, ByVal SelMonth As Integer,
                                              ByVal SelYear As Integer, ByVal CostTypeVal As Integer, ByVal dummyTableName As String) As Integer

        Dim strSQL As String
        Dim LastSelYear As Integer
        Dim LastSelMonth As Integer
        Dim DateTimeString As String = "{ d '" + SelYear.ToString + "-" + SelMonth.ToString + "-1' }"
        Dim LastMonthString As String = "{ d '" + LastSelYear.ToString + "-" + LastSelMonth.ToString + "-1' }"
        Dim TestString As String = ""
        Dim IDString As StringBuilder = New StringBuilder
        Dim i As Integer
        Dim ChkData As DataTable
        Dim ChkLastMonth As DataTable
        Dim Exist As Boolean = False
        Dim ChkLink As DataTable
        Dim CostGroupID As Integer = 0
        Dim LastMonthCostGroupID As Integer = 0

        If SelMonth = 1 Then
            LastSelMonth = 12
            LastSelYear = SelYear - 1
        Else
            LastSelMonth = SelMonth - 1
            LastSelYear = SelYear
        End If

        If CostTypeVal = -2 Then
            If Exist = True Then
                ChkLink = dbUtil.List("select * from MaterialCostGroupLinkInventory", connection)
            Else
                ChkLink = dbUtil.List("select * from Property where 0=1", connection)
            End If
            If ChkLink.Rows.Count = 0 Then
                strSQL = "select * from MaterialCostGroup where StartDate <= " + DateTimeString + " AND EndDate > " + DateTimeString + " order by StartDate"
                TestString += "<br>" + strSQL
                ChkData = dbUtil.List(strSQL, connection)

                strSQL = "select * from MaterialCostGroup where StartDate <= " + LastMonthString + " AND EndDate > " + LastMonthString + " order by StartDate"
                TestString += "<br>" + strSQL

                ChkLastMonth = dbUtil.List(strSQL, connection)
            Else
                strSQL = "select a.* from MaterialCostGroup a, MaterialCostGroupLinkInventory b where a.MaterialCostGroupID=b.MaterialCostGroupID AND b.InventoryID=" + ShopID.ToString + " AND StartDate <= " + DateTimeString + " AND EndDate > " + DateTimeString + " order by StartDate"
                TestString += "<br>" + strSQL
                ChkData = dbUtil.List(strSQL, connection)

                strSQL = "select a.* from MaterialCostGroup a, MaterialCostGroupLinkInventory b where a.MaterialCostGroupID=b.MaterialCostGroupID AND b.InventoryID=" + ShopID.ToString + " AND StartDate <= " + LastMonthString + " AND EndDate > " + LastMonthString + " order by StartDate"
                TestString += "<br>" + strSQL

                ChkLastMonth = dbUtil.List(strSQL, connection)
            End If

            If ChkData.Rows.Count > 0 Then CostGroupID = ChkData.Rows(0)("MaterialCostGroupID")
            If ChkLastMonth.Rows.Count > 0 Then LastMonthCostGroupID = ChkLastMonth.Rows(0)("MaterialCostGroupID")

            dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "', 'U') IS NOT NULL DROP TABLE " + dummyTableName, connection)
            dbUtil.sqlExecute("create table " + dummyTableName + " (MaterialID int, TotalPrice decimal(18,4), TotalAmount decimal(18,4), BeginningPricePerUnit decimal(18,4), BeginningAmount decimal(18,4), PricePerUnit decimal(18,4), RecTotalPrice decimal(18,4), RecTotalAmount decimal(18,4), PRIMARY KEY (MaterialID))", connection)

            strSQL = " select a.MaterialID,CASE WHEN b.MaterialPrice is NULL THEN 0 ELSE b.MaterialPrice END AS TotalPrice" + _
                " ,CASE WHEN b.UnitSmallRatio is NULLTHEN 1 ELSE b.UnitSmallRatio END AS TotalAmount,CASE WHEN c.MaterialPrice is NULLTHEN 0 ELSE c.MaterialPrice END AS BeginningPricePerUnit, CASE WHEN c.UnitSmallRatio is NULLTHEN 1 ELSE c.UnitSmallRatio END As BeginningAmount,1,CASE WHEN b.MaterialPrice is NULLTHEN 0 ELSE b.MaterialPrice END AS RecTotalPrice,CASE WHEN b.UnitSmallRatio is NULLTHEN 1 ELSE b.UnitSmallRatio END AS RecTotalAmount from Materials a left outer join MaterialCostTable b ON a.MaterialID=b.MaterialID AND b.MaterialCostGroupID=" + CostGroupID.ToString + " left outer join MaterialCostTable c ON a.MaterialID=c.MaterialID AND c.MaterialCostGroupID=" + LastMonthCostGroupID.ToString
            dbUtil.sqlExecute("insert into " + dummyTableName + strSQL, connection)
            TestString += "<br>" + strSQL
        ElseIf CostTypeVal = -1 Then
            dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "', 'U') IS NOT NULL DROP TABLE " + dummyTableName, connection)
            dbUtil.sqlExecute("create table " + dummyTableName + " (MaterialID int, TotalPrice decimal(18,4), TotalAmount decimal(18,4), BeginningPricePerUnit decimal(18,4), BeginningAmount decimal(18,4), PricePerUnit decimal(18,4), RecTotalPrice decimal(18,4), RecTotalAmount decimal(18,4), PRIMARY KEY (MaterialID))", connection)
        Else
            ' Get standard cost in case amount is 0 from beginning stock _Begin
            strSQL = "select b.ProductID AS MaterialID,b.ProductPricePerUnit from document a, docdetail b, documenttype c where a.documentid=b.documentid and a.shopid=b.shopid and a.documenttypeid=c.documenttypeid and a.shopid=c.shopid and c.langid=1 and a.documenttypeid=10 and a.documentstatus=2 and MONTH(a.documentdate) = " + SelMonth.ToString + " AND YEAR(a.documentdate) = " + SelYear.ToString + " and a.productlevelid=" + ShopID.ToString
            dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "_Begin', 'U') IS NOT NULL DROP TABLE " + dummyTableName + "_Begin", connection)
            dbUtil.sqlExecute("create table " + dummyTableName + "_Begin (MaterialID int, BeginningPricePerUnit decimal(18,4), PRIMARY KEY (MaterialID))", connection)
            dbUtil.sqlExecute("insert into " + dummyTableName + "_Begin " + strSQL, connection)

            Dim dtTable As DataTable = dbUtil.List("select * from " + dummyTableName + "_Begin", connection)

            For i = 0 To dtTable.Rows.Count - 1
                IDString = IDString.Append(dtTable.Rows(i)("MaterialID").ToString + ",")
            Next
            IDString = IDString.Append("0")
            Dim IDList As String = IDString.ToString

            If Trim(IDList) = "" Then IDList = "0"
            strSQL = " select b.MaterialID,a.BeginningPricePerUnit from " + dummyTableName + "_Begin a left outer join Materials b ON a.MaterialID=b.MaterialID where b.MaterialID is not NULL AND b.MaterialID NOT IN (" + IDList + ")"
            dbUtil.sqlExecute("insert into " + dummyTableName + "_Begin " + strSQL, connection)

            '----------- Calculate Receive Average Cost  _Rec
            dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "_Rec', 'U') IS NOT NULL DROP TABLE " + dummyTableName + "_Rec", connection)
            dbUtil.sqlExecute("create table " + dummyTableName + "_Rec (MaterialID int, RecTotalPrice decimal(18,4), RecTotalAmount decimal(18,4), PRIMARY KEY (MaterialID))", connection)

            strSQL = " select b.ProductID as MaterialID, sum(c.MovementInStock*b.ProductNetPrice) as totalPrice, sum(c.MovementInStock*b.UnitSmallAmount) as totalAmount from document a, docdetail b, documenttype c where a.documentid=b.documentid and a.shopid=b.shopid and a.documenttypeid=c.documenttypeid and a.shopid=c.shopid and c.langid=1 and c.calculatestandardprofitloss=1 and a.documentstatus=2 and c.DocumentTypeID IN (2,39) and MONTH(a.documentdate) = " + SelMonth.ToString + " AND YEAR(a.documentdate) = " + SelYear.ToString + " and a.productlevelid=" + ShopID.ToString + " group by b.ProductID"
            dbUtil.sqlExecute("insert into " + dummyTableName + "_Rec " + strSQL, connection)

            '----------- End Rec Cal ----------
            strSQL = " select b.ProductID as MaterialID, sum(c.MovementInStock*b.ProductNetPrice) as totalPrice, sum(c.MovementInStock*b.UnitSmallAmount) as totalAmount,CASE WHEN d.BeginningPricePerUnit is NULL THEN 0 ELSE d.BeginningPricePerUnit END AS BeginningPricePerUnit from document a inner join docdetail b ON a.documentid=b.documentid and a.shopid=b.shopid inner join documenttype c ON a.documenttypeid=c.documenttypeid and a.shopid=c.shopid left outer join " + dummyTableName + "_Begin d ON b.ProductID=d.MaterialID where  c.langid=1 and c.calculatestandardprofitloss=1 and a.documentstatus=2 and MONTH(a.documentdate) = " + SelMonth.ToString + " AND YEAR(a.documentdate) = " + SelYear.ToString + " and a.productlevelid=" + ShopID.ToString + " group by  b.ProductID,d.BeginningPricePerUnit "
            '_Stock
            dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "_Stock', 'U') IS NOT NULL DROP TABLE " + dummyTableName + "_Stock", connection)
            dbUtil.sqlExecute("create table " + dummyTableName + "_Stock (MaterialID int, TotalPrice decimal(18,4), TotalAmount decimal(18,4), BeginningPricePerUnit decimal(18,4), PRIMARY KEY (MaterialID))", connection)
            dbUtil.sqlExecute("insert into " + dummyTableName + "_Stock " + strSQL, connection)

            strSQL = " select a.MaterialID,TotalPrice,TotalAmount,BeginningPricePerUnit,CASE WHEN TotalAmount = 0 THEN BeginningPricePerUnit ELSE TotalPrice/TotalAmount END AS PricePerUnit,1, RecTotalPrice, RecTotalAmount from " + dummyTableName + "_Stock a left outer join " + dummyTableName + "_Rec b ON a.MaterialID=b.MaterialID"
            dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "', 'U') IS NOT NULL DROP TABLE " + dummyTableName, connection)
            dbUtil.sqlExecute("create table " + dummyTableName + " (MaterialID int, TotalPrice decimal(18,4), TotalAmount decimal(18,4), BeginningPricePerUnit decimal(18,4), BeginningAmount decimal(18,4), PricePerUnit decimal(18,4), RecTotalPrice decimal(18,4), RecTotalAmount decimal(18,4), PRIMARY KEY (MaterialID))", connection)
            dbUtil.sqlExecute("insert into " + dummyTableName + strSQL, connection)

            dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "_Stock', 'U') IS NOT NULL DROP TABLE " + dummyTableName + "_Stock", connection)
            dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "_Begin', 'U') IS NOT NULL DROP TABLE " + dummyTableName + "_Begin", connection)
            dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "_Rec', 'U') IS NOT NULL DROP TABLE " + dummyTableName + "_Rec", connection)

        End If

    End Function

    Friend Function CheckMaterialStockBelowZero(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal startDate As Date, ByVal endDate As Date, ByVal shopId As Integer,
                                                ByVal notIncludeMaterialList As String, ByVal dummyTableName As String) As DataTable

        Dim dtDateBeginStock As New DataTable
        Dim dtStockData As New DataTable
        Dim strSQL As String = ""
        Dim dateBegin As Date
        Dim strCurrentDate As String = FormatDateTime(Date.Now)
        Dim dtResult As New DataTable

        dummyTableName = "Dummy_" & dummyTableName
        strSQL = "IF OBJECT_ID('" & dummyTableName & "', 'U') IS NOT NULL DROP TABLE " & dummyTableName & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Create Table " & dummyTableName & " (MaterialId int NOT NULL, MaterialCode nvarchar(50) NULL, MaterialName nvarchar(100), Qty int NOT NULL DEFAULT '0',UnitSmallID int NOT NULL DEFAULT '0'); "
        dbUtil.sqlExecute(strSQL, connection)

        dtDateBeginStock = dbUtil.List("SELECT MAX(documentdate) AS maxtransferdate FROM document WHERE documenttypeid=10 AND productlevelid=" & shopId & " AND documentstatus=2", connection)
        If dtDateBeginStock.Rows.Count > 0 Then
            If Not IsDBNull(dtDateBeginStock.Rows(0)("maxtransferdate")) Then
                dateBegin = dtDateBeginStock.Rows(0)("maxtransferdate")
            Else
                dateBegin = startDate
            End If
        Else
            dateBegin = startDate
        End If
        strSQL = "INSERT INTO " & dummyTableName & "(MaterialId, MaterialCode, MaterialName, Qty, UnitSmallID)" & _
              " SELECT dd.productid, m.materialcode, m.materialname, SUM(dd.unitsmallamount * dt.movementinstock) AS qty,m.UnitSmallID" &
              " FROM document d, docdetail dd, documenttype dt, materials m" & _
              " WHERE d.documentid=dd.documentid AND d.shopid=dd.shopid AND " & _
              " d.documentdate>=" & FormatDate(dateBegin) & " AND d.documentdate<=" & FormatDate(endDate) & _
              " AND d.documentstatus=2 AND d.productlevelid= " & shopId & " AND dt.shopid=d.shopid " & _
              " AND dt.documenttypeid = d.documenttypeid And dt.langid = 2 " & _
              " AND dd.productid = m.materialid" & _
              " AND dd.productid NOT IN(" & notIncludeMaterialList & ")" & _
              " GROUP BY dd.productid, m.materialcode, m.materialname,m.UnitSmallID"
        dbUtil.sqlExecute(strSQL, connection)

        strSQL = "SELECT * FROM " & dummyTableName & " WHERE Qty < 0"
        dtResult = dbUtil.List(strSQL, connection)
        dbUtil.sqlExecute("IF OBJECT_ID('" & dummyTableName & "', 'U') IS NOT NULL DROP TABLE " & dummyTableName & ";", connection)
        Return dtResult
    End Function

    Friend Function CheckMonthlyDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal selMonth As Integer, ByVal selYear As Integer, ByVal inventoryId As Integer) As DataTable
        Dim strSQL As String = ""
        strSQL = "select * from document where documenttypeid=7  and documentstatus<>99 and DocumentMonth=" & selMonth & " and documentyear=" & selYear & " and ProductLevelID=" & inventoryId
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function DeleteZeroCompareAmountMaterialInDocDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction,
                                                                      ByVal documentId As Integer, ByVal documentShopId As Integer) As Integer
        Dim strSQL As String
        strSQL = "Delete From DocDetail " & _
                 "Where DocumentID = " & documentId & " AND ShopID = " & documentShopId & " AND " & _
                 " UnitSmallAmount = 0 AND ProductAmount = 0  "
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function DeleteDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                                ByVal documentShopId As Integer, ByVal docDetailID As Integer) As Integer
        Dim strSQL As String
        strSQL = "Delete From DocDetail " &
                 "Where DocDetailID = " & docDetailID & " AND DocumentID = " & documentId & " AND ShopID = " & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function DeleteDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                               ByVal documentShopId As Integer, ByVal strDelDocID As String) As Integer
        Dim strSQL As String
        strSQL = "Delete From DocDetail " &
                 "Where  DocumentID = " & documentId & " AND ShopID = " & documentShopId & " AND DocDetailID IN (" & Mid(strDelDocID, 1, Len(strDelDocID) - 2) & ")"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function DocumentIsAlreadyReferTo(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction,
                                                              ByVal refDocumentID As Integer, ByVal refDocumentShopID As Integer, ByVal documentType As Integer,
                                                              notDocumentID As Integer, notDocumentShopID As Integer) As Boolean
        Dim strSQL As String
        Dim dtResult As DataTable
        strSQL = "Select DocumentID From Document " & _
                 "Where DocumentIDRef = " & refDocumentID & " AND DocumentIDRefShopID = " & refDocumentShopID & " AND " & _
                 " DocumentStatus = " & GlobalVariable.DOCUMENTSTATUS_APPROVE & " AND DocumentTypeID = " & documentType & " AND " & _
                 " NOT (DocumentID = " & notDocumentID & " AND ShopID = " & notDocumentShopID & ") "
        dtResult = dbUtil.List(strSQL, connection, objTrans)
        If dtResult.Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Friend Function DeleteMaterialInTableDefaultPrice(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction,
                                                      ByVal inventoryID As Integer, ByVal vendorId As Integer, ByVal strMaterialId As String) As Integer
        Dim strSQL As String
        strMaterialId = Mid(strMaterialId, 1, Len(strMaterialId) - 2)
        strSQL = "Delete From MaterialDefaultPrice " &
                 "Where MaterialID IN (" & strMaterialId & ") AND InventoryID = " & inventoryID & " AND VendorID = " & vendorId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Sub DropTempTable(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal dummyTableName As String)
        Dim strSQL As String = ""
        dbUtil.sqlExecute("IF OBJECT_ID('" + dummyTableName + "', 'U') IS NOT NULL DROP TABLE " + dummyTableName, connection)
        dbUtil.sqlExecute("IF OBJECT_ID('DummyStockCard', 'U') IS NOT NULL DROP TABLE DummyStockCard", connection)
    End Sub

    Friend Function FinishDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal docID As Integer,
                                          ByVal docShopID As Integer, ByVal updateDate As String) As Integer
        Dim strSQL As String
        strSQL = "Update Document " &
                 "Set DocumentStatus = 4, UpdateDate = " & updateDate & ", AlreadyExportToHQ = 0, AlreadyExportToBranch = 0 " &
                 "Where DocumentID = " & docID & " AND ShopID = " & docShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function GetDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentId As Integer,
                                       ByVal documentShopId As Integer, ByVal langID As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select d.*, dt.DocumentTypeHeader, dt.DocumentTypeName,dt.MovementInStock" &
                 " From Document d, DocumentType dt " &
                 " Where d.DocumentID = " & documentId &
                 " AND d.ShopID = " & documentShopId &
                 " AND d.ShopID = dt.ShopID AND  d.DocumentTypeID = dt.DocumentTypeID " &
                 " AND dt.LangID = " & langID & " AND dt.ComputerID = 0 "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                      ByVal documentShopId As Integer, ByVal langID As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select d.*, dt.DocumentTypeHeader, dt.DocumentTypeName,dt.MovementInStock" &
                 " From Document d, DocumentType dt " &
                 " Where d.DocumentID = " & documentId &
                 " AND d.ShopID = " & documentShopId &
                 " AND d.ShopID = dt.ShopID AND  d.DocumentTypeID = dt.DocumentTypeID " &
                 " AND dt.LangID = " & langID & " AND dt.ComputerID = 0 "
        Return dbUtil.List(strSQL, connection, objTrans)
    End Function

    Friend Function GetSaleDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentDate As String, ByVal inventoryId As Integer, ByVal langID As Integer, ByVal documentTypeId As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select d.*, dt.DocumentTypeHeader, dt.DocumentTypeName,dt.MovementInStock" &
                 " From Document d, DocumentType dt " &
                 " Where d.DocumentDate = " & documentDate &
                 " AND d.ShopID = " & inventoryId &
                 " AND d.ShopID = dt.ShopID AND  d.DocumentTypeID = dt.DocumentTypeID AND d.documentstatus=2 " &
                 " AND dt.LangID = " & langID & " AND dt.ComputerID = 0 " &
                 " AND d.DocumentTypeID=" & documentTypeId
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetDocumentWorkingProcess(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentDate As String, ByVal documentTypeId As Integer,
                                              ByVal shopid As Integer, ByVal langID As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select d.*, dt.DocumentTypeHeader, dt.DocumentTypeName,dt.MovementInStock" &
                 " From Document d, DocumentType dt " &
                 " Where d.DocumentDate = " & documentDate &
                 " AND d.ShopID = " & shopid &
                 " AND d.DocumentTypeId = " & documentTypeId &
                 " AND d.DocumentStatus = 1" &
                 " AND d.ShopID = dt.ShopID AND  d.DocumentTypeID = dt.DocumentTypeID " &
                 " AND dt.LangID = " & langID & " AND dt.ComputerID = 0 "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetUser(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection) As DataTable
        Dim strSQL As String
        strSQL = "select USER_ID as staffId, USERDESC as StaffName, USERNAME,PASSWORD from TBUSER"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetMaxDocumentNumber(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentShopId As Integer,
                                                ByVal documentTypeId As Integer, ByVal docYear As Integer, ByVal docMonth As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select MaxDocumentNumber " & _
                "From MaxDocumentNumber " & _
                "Where DocumentYear = " & docYear & " AND DocumentMonth = " & docMonth & " AND " & _
                " DocType = " & documentTypeId & " AND ShopID = " & documentShopId
        Return dbUtil.List(strSQL, connection, objTrans)
    End Function

    Friend Function GetMaxDocumentNumberFromDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentShopId As Integer,
                                                ByVal documentTypeId As Integer, ByVal docYear As Integer, ByVal docMonth As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select Count(*) as NumRow, Max(DocumentNumber) + 1 as NewID " & _
                        "From Document " & _
                        "Where DocumentYear = " & docYear & " AND DocumentMonth = " & docMonth & " AND " & _
                        " ShopID = " & documentShopId & " AND DocumentTypeID = " & documentTypeId
        Return dbUtil.List(strSQL, connection, objTrans)
    End Function

    Friend Function GetDocumentNumber(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentId As Integer, ByVal documentShopId As Integer,
                                                   ByVal langID As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select d.DocumentYear, d.DocumentMonth, d.DocumentNumber, dt.DocumentTypeHeader, d.DocumentStatus " & _
                "From Document d, DocumentType dt " & _
                "Where d.DocumentID = " & documentId & " AND d.ShopID = " & documentShopId & " AND " & _
                " d.DocumentTypeID = dt.DocumentTypeID AND dt.LangID = " & langID & " AND dt.ComputerID = 0 AND " & _
                " d.ShopID = dt.ShopID "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentId As Integer, ByVal documentShopId As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select dd.*, m.MaterialID, m.MaterialCode, m.MaterialName, m.UnitSmallID, m.MaterialTaxType " & _
                 "From DocDetail dd, Materials m " & _
                 "Where dd.DocumentID = " & documentId & " AND dd.ShopID = " & documentShopId & " AND dd.ProductID = m.MaterialID " & _
                 "Order by dd.DocumentID, dd.ShopID, dd.DocDetailID "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetDocumentDetailByMaterialID(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentId As Integer, ByVal documentShopId As Integer, ByVal materialId As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select dd.*, m.MaterialID, m.MaterialCode, m.MaterialName, m.UnitSmallID, m.MaterialTaxType " & _
                 "From DocDetail dd, Materials m " & _
                 "Where dd.DocumentID = " & documentId & " AND dd.ShopID = " & documentShopId & " AND dd.ProductID = m.MaterialID AND dd.ProductID=" & materialId & _
                 "Order by dd.DocumentID, dd.ShopID, dd.DocDetailID "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer, ByVal documentShopId As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select dd.*, m.MaterialID, m.MaterialCode, m.MaterialName, m.UnitSmallID, m.MaterialTaxType " & _
                 "From DocDetail dd, Materials m " & _
                 "Where dd.DocumentID = " & documentId & " AND dd.ShopID = " & documentShopId & " AND dd.ProductID = m.MaterialID " & _
                 "Order by dd.DocumentID, dd.ShopID, dd.DocDetailID "
        Return dbUtil.List(strSQL, connection, objTrans)
    End Function

    Friend Function GetLastTransferStock(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal inventoryId As Integer) As DataTable
        Dim strSQL As String
        strSQL = "select top 1 documentdate from document where documenttypeid=10 and ProductLevelID=" & inventoryId & " order by documentdate desc"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetMaxDocumentID(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal docShopID As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select MaxDocumentID as MaxDocID From MaxDocumentID " &
                 " Where ShopID = " & docShopID & " AND IsDocumentOrBatch = 0"
        Return dbUtil.List(strSQL, connection, objTrans)
    End Function

    Friend Function GetMaxDocumentIDFromTableDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal docShopID As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select Max(DocumentID) as MaxDocID, Count(DocumentID) as CountRow " &
                    "From Document " &
                    "Where ShopID = " & docShopID
        Return dbUtil.List(strSQL, connection, objTrans)
    End Function

    Friend Function GetDocumentTypeStockCountMaterialSetting(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal productLevelID As Integer, ByVal strDocType As String) As DataTable
        Dim strSQL As String
        strSQL = "Select Distinct DocumentTypeID From StockCountMaterialSetting Where DocumentTypeID NOT IN (0," & strDocType & ") "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetDocumentTypeRedue(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal shopID As Integer, ByVal documentTypeID As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select * From DocumentType Where movementinstock=-1 AND DocumentTypeID = " & documentTypeID & " AND ShopID = " & shopID & " AND langid = 2"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetLastApproveDocumentDate(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal productLevelID As Integer, ByVal strDocType As String) As DataTable
        Dim strSQL As String
        strSQL = "Select Max(DocumentDate) as DocumentDate, Max(ApproveDate) as ApproveDate " & _
                 "From Document " & _
                 "Where DocumentTypeID In (" & strDocType & ") AND " & _
                " ProductLevelID = " & productLevelID & " AND DocumentStatus = " & GlobalVariable.DOCUMENTSTATUS_APPROVE
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetLastTransferStockOrCountStockDocumentDate(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal productLevelID As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select Max(DocumentDate) as DocumentDate " & _
                 "From Document " & _
                 "Where DocumentTypeID IN (" & GlobalVariable.DOCUMENTTYPE_TRANSFERSTOCK & "," & GlobalVariable.DOCUMENTTYPE_WEEKLYSTOCK & "," & GlobalVariable.DOCUMENTTYPE_DAILYSTOCK & ") AND ProductLevelID = " & productLevelID & " AND DocumentStatus = 2 "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function CheckCountStock(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal inventoryId As Integer, ByVal selEndDate As String) As Boolean
        Dim strSQL As String
        Dim dt As New DataTable
        strSQL = "Select * From Document " & _
                 "Where DocumentTypeID IN (" & GlobalVariable.DOCUMENTTYPE_TRANSFERSTOCK & "," & GlobalVariable.DOCUMENTTYPE_WEEKLYSTOCK & "," & GlobalVariable.DOCUMENTTYPE_DAILYSTOCK & ") AND ProductLevelID = " & inventoryId & " AND DocumentStatus = 2 AND DocumentDate=" & selEndDate
        dt = dbUtil.List(strSQL, connection)
        If dt.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Friend Function GetDocumentFromReference(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal refDocumentID As Integer, ByVal refDocumentShopID As Integer,
                                                    ByVal documentType As String, ByVal documentStatus As String) As DataTable
        Dim strSQL As String
        strSQL = "Select DocumentID, ShopID, ProductLevelID, FromInvID, ToInvID, DocumentTypeID, DocumentStatus From Document " & _
                 "Where DocumentIDRef = " & refDocumentID & " AND DocumentIDRefShopID = " & refDocumentShopID & " AND " & _
                 " DocumentTypeID IN (" & documentType & ") "
        If documentStatus <> "" Then
            strSQL &= " AND DocumentStatus IN (" & documentStatus & ") "
        End If
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetMaterialFromOriginalDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer, ByVal documentShopId As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select dd.*,0.00 as UnitRatio, dd.UnitSmallAmount as OriginalSmallAmount " & _
                 " From DocDetail dd Where DocumentID = " & documentId & " AND ShopID = " & documentShopId & _
                 " Order By DocDetailID "
        Return dbUtil.List(strSQL, connection, objTrans)
    End Function

    Friend Function GetMaterialFromReferenceDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentRefId As Integer, ByVal documentRefShopId As Integer,
                                                     ByVal notIncludeDocId As Integer, ByVal notIncludeDocShopId As Integer, ByVal productIdList As String, ByVal documentTypeId As Integer) As DataTable
        Dim strSQL As String
        strSQL = " Select dd.*, 0 as AlreadyProcess " &
                "From Document d, DocDetail dd " &
                "Where d.DocumentIDRef = " & documentRefId & " AND d.DocumentIDRefShopID = " & documentRefShopId & " AND " & _
                " d.DocumentID = dd.DocumentID AND dd.ShopID = d.ShopID AND d.DocumentStatus NOT IN (0,99) AND " & _
                " dd.ProductID IN (" & productIdList & ") AND NOT (d.documentID = " & notIncludeDocId & " AND d.ShopID =" & notIncludeDocShopId & " ) " & _
                "Order By dd.ShopID, dd.DocumentID, dd.DocDetailID "
        Return dbUtil.List(strSQL, connection, objTrans)
    End Function

    Friend Function GetMaxDocDetailID(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                             ByVal documentShopId As Integer) As Integer
        Dim strSQL As String = ""
        Dim newDocDetailID As Integer = 0
        Dim dtResult As New DataTable

        strSQL = "Select Max(DocDetailID) As MaxDocDetailID From DocDetail " &
                 "Where DocumentId = " & documentId & " And ShopId = " & documentShopId
        dtResult = dbUtil.List(strSQL, connection, objTrans)
        If dtResult.Rows.Count > 0 Then
            If Not IsDBNull(dtResult.Rows(0)("MaxDocDetailID")) Then
                newDocDetailID = dtResult.Rows(0)("MaxDocDetailID") + 1
            Else
                newDocDetailID = 1
            End If
        Else
            newDocDetailID = 1
        End If
        Return newDocDetailID
    End Function

    Friend Function GetMaterialForUpdateDefaultPrice(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction,
                                                            ByVal documentID As Integer, ByVal documentShopID As Integer) As DataTable

        Dim strSQL As String
        strSQL = "Select m.MaterialID, dd.ProductPricePerUnit, dd.UnitSmallAmount, dd.ProductAmount, ur.UnitSmallRatio, " &
                     " ur.UnitSmallID, ur.UnitLargeID " &
                     " From DocDetail dd, Materials m, UnitRatio ur " &
                     " Where dd.ProductID = m.MaterialID AND m.AutoUpdateMaterialPrice = 1 AND " &
                     " dd.UnitID = ur.UnitLargeID AND dd.ProductUnit = ur.UnitSmallID AND " &
                     " dd.DocumentID = " & documentID & " AND dd.ShopID = " & documentShopID & " " &
                     " Order by MaterialID, ur.UnitLargeID, UnitSmallRatio "
        Return dbUtil.List(strSQL, connection, objTrans)
    End Function

    Friend Function GetAddReduceDocumentType(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal movementInStock As Integer,
                                             ByVal langID As Integer, ByVal shopID As Integer) As DataTable
        Dim strSQL As String
        Dim strMovement As String
        If movementInStock = 0 Then
            strMovement = ""
        Else
            strMovement = " AND MovementInStock = " & movementInStock & " "
        End If
        strSQL = "Select * " & _
                "From DocumentType " & _
                "Where IsAddReduceDoc = 1 AND LangID = " & langID & " AND ShopID = " & shopID & " AND Deleted = 0 AND " & _
                " ShowOnSearch = 1 " & strMovement & _
                "Order By DocumentTypeID "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetDocumentIDRef(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentTypeId As Integer,
                                     ByVal refDocumentID As Integer, ByVal refDocumentShopID As Integer) As DataTable
        Dim strSQL As String = ""
        strSQL = "select * from Document where DocumentTypeID=" & documentTypeId &
                 " And DocumentIDRef=" & refDocumentID &
                 " And DocumentIDRefShopID=" & refDocumentShopID &
                 " And DocumentStatus =1"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetMaterialStock(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal materialGroupId As Integer, ByVal materialDeptId As Integer, ByVal materialCode As String,
                                     ByVal dummyMaterialStandardCostTable As String, ByRef groupData As DataTable, ByRef groupDataL As DataTable) As DataTable

        Dim strSQL As String = ""
        Dim selectString As String
        Dim whereString As String
        Dim additionalQuery As String = ""

        groupData = dbUtil.List("select * from DocumentTypeGroup where Ordering > 0 order by Ordering", connection)
        selectString = "u.UnitSmallName,a.UnitSmallID,a.MaterialID,a.MaterialName,a.MaterialCode,std.TotalPrice,std.TotalAmount,std.BeginningPricePerUnit,std.RecTotalPrice,std.RecTotalAmount,b.CalculateStandardProfitLoss AS CalculateStandardProfitLoss0,b.NetSmallAmount AS NetSmallAmount0,b.TotalAmount AS TotalAmount0,b.ProductNetPrice AS ProductNetPrice0"
        whereString = " left outer join UnitSmall u ON a.UnitSmallID=u.UnitSmallID left outer join " & dummyMaterialStandardCostTable & " std ON a.MaterialID=std.MaterialID left outer join DummyStockCard b on a.MaterialID=b.MaterialID AND b.Ordering=0 "

        Dim i As Integer
        For i = 0 To groupData.Rows.Count - 1
            selectString += ",b" + groupData.Rows(i)("DocumentTypeGroupID").ToString + ".CalculateStandardProfitLoss AS CalculateStandardProfitLoss" + groupData.Rows(i)("DocumentTypeGroupID").ToString +
                            ",b" + groupData.Rows(i)("DocumentTypeGroupID").ToString + ".NetSmallAmount As NetSmallAmount" + groupData.Rows(i)("DocumentTypeGroupID").ToString +
                            ",b" + groupData.Rows(i)("DocumentTypeGroupID").ToString +
                            ".TotalAmount AS TotalAmount" + groupData.Rows(i)("DocumentTypeGroupID").ToString +
                            ",b" + groupData.Rows(i)("DocumentTypeGroupID").ToString +
                            ".ProductNetPrice AS ProductNetPrice" + groupData.Rows(i)("DocumentTypeGroupID").ToString
            whereString += " left outer join DummyStockCard b" + groupData.Rows(i)("DocumentTypeGroupID").ToString +
                           " on a.MaterialID=b" + groupData.Rows(i)("DocumentTypeGroupID").ToString +
                           ".MaterialID AND b" + groupData.Rows(i)("DocumentTypeGroupID").ToString +
                           ".DocumentTypeGroupID=" + groupData.Rows(i)("DocumentTypeGroupID").ToString
        Next
        groupDataL = dbUtil.List("select * from DocumentTypeGroup where Ordering < 0 order by Ordering", connection)
        For i = 0 To groupDataL.Rows.Count - 1
            selectString += ",b" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                            ".CalculateStandardProfitLoss AS CalculateStandardProfitLoss" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                            ",b" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                            ".NetSmallAmount As NetSmallAmount" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                            ",b" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                            ".TotalAmount AS TotalAmount" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                            ",b" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                            ".ProductNetPrice AS ProductNetPrice" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString
            whereString += " left outer join DummyStockCard b" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                           " on a.MaterialID=b" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                           ".MaterialID AND b" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString +
                           ".DocumentTypeGroupID=" + groupDataL.Rows(i)("DocumentTypeGroupID").ToString
        Next

        If materialGroupId > 0 Then
            additionalQuery += " AND mg.MaterialGroupId =" & materialGroupId
        End If
        If materialDeptId > 0 Then
            additionalQuery += " AND md.MaterialDeptId =" & materialDeptId
        End If
        If materialCode <> "" Then
            additionalQuery += " AND a.MaterialCode Like '%" & ReplaceSuitableStringForSQL(materialCode) & "%'"
        End If
        strSQL = "select " + selectString + ",mg.MaterialGroupID,mg.MaterialGroupCode,mg.MaterialGroupName from Materials a left outer join MaterialDept md ON a.MaterialDeptID=md.MaterialDeptID " +
                 "left outer join MaterialGroup mg ON md.MaterialGroupID=mg.MaterialGroupID " + whereString +
                 " where a.deleted=0 and a.MaterialIDRef=0 " + additionalQuery
        'DocumentSQL.InsertLog(dbUtil, connection, "StockCard", "SearchByCode", "77", strSQL.ToString)
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetCurrentStockForResetStockToZero(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentDate As Date, ByVal shopID As Integer,
                                                       ByVal staffID As Integer, ByVal stockCountDocType As Integer) As DataTable
        Dim strSQL, strStartDate, strEndDate As String
        Dim startDate As Date
        Dim dtResult As DataTable
        startDate = documentDate.AddDays(1 - documentDate.Day)
        strStartDate = FormatDate(startDate)
        strEndDate = FormatDate(documentDate)

        'All Material In Stock Count
        strSQL = "IF OBJECT_ID('AllMaterialInDailyStock', 'U') IS NOT NULL DROP TABLE AllMaterialInDailyStock" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Create Table AllMaterialInDailyStock" & staffID & " (MaterialID int NOT NULL); "
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Insert INTO AllMaterialInDailyStock" & staffID & "(MaterialID) " & _
                "Select Distinct MaterialID From Materials Where Deleted=0"
        dbUtil.sqlExecute(strSQL, connection)

        strSQL = "IF OBJECT_ID('CurrentMaterialInDailyStock', 'U') IS NOT NULL DROP TABLE CurrentMaterialInDailyStock" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "Create Table CurrentMaterialInDailyStock" & staffID & _
                " (MaterialID int NOT NULL, CurrentAmount decimal(18,4)); "
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "ALTER TABLE CurrentMaterialInDailyStock" & staffID & " ADD INDEX MaterialInStockIndex (MaterialID);"
        dbUtil.sqlExecute(strSQL, connection)
        'Insert CurrentMaterial That has in Material In Stock
        strSQL = "Insert INTO CurrentMaterialInDailyStock" & staffID & "(MaterialID, CurrentAmount) " & _
                "Select dd.ProductID, Sum(dt.MovementInStock * dd.UnitSmallAmount) " & _
                 "From Document d, DocDetail dd, DocumentType dt, AllMaterialInDailyStock" & staffID & " a " & _
                "Where d.DocumentID = dd.DocumentID AND d.ShopID = dd.ShopID AND d.DocumentDate >= " & strStartDate & " AND " & _
                " d.DocumentDate <= " & strEndDate & " AND d.DocumentStatus = " & GlobalVariable.DOCUMENTSTATUS_APPROVE & " AND " & _
                " d.ProductLevelID = " & shopID & " AND d.DocumentTypeID = dt.DocumentTypeID AND d.ShopID = dt.ShopID AND " & _
                " dt.LangID = 2 AND dt.ComputerID = 0 AND dd.ProductID = a.MaterialID " & _
                " Group By dd.ProductID "
        dbUtil.sqlExecute(strSQL, connection)

        'Insert CurrentMaterial That has in Material In Stock
        strSQL = "Insert INTO CurrentMaterialInDailyStock" & staffID & "(MaterialID, CurrentAmount) " & _
                "Select a.MaterialID, 0 " & _
                 "From AllMaterialInDailyStock" & staffID & " a LEFT OUTER JOIN CurrentMaterialInDailyStock" & staffID & " c ON " & _
                 " a.MaterialID = c.MaterialID " & _
                 "Where c.MaterialID IS NULL "
        dbUtil.sqlExecute(strSQL, connection)

        strSQL = "Select c.*, ur.UnitSmallID, ur.UnitLargeID, ul.UnitLargeName " & _
                 "From CurrentMaterialInDailyStock" & staffID & " c, Materials m, UnitRatio ur, UnitLarge ul " & _
                 "Where c.MaterialID = m.MaterialID AND m.UnitSmallID = ur.UnitSmallID AND ur.Deleted = 0 AND ur.UnitLargeID = ul.UnitLargeID AND " & _
                 " ur.UnitLargeRatio = 1 AND ur.UnitSmallRatio = 1 " & _
                 "Order By c.MaterialID, ur.UnitLargeID "
        dtResult = dbUtil.List(strSQL, connection)

        strSQL = "IF OBJECT_ID('CurrentMaterialInDailyStock', 'U') IS NOT NULL DROP TABLE CurrentMaterialInDailyStock" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        strSQL = "IF OBJECT_ID('AllMaterialInDailyStock', 'U') IS NOT NULL DROP TABLE AllMaterialInDailyStock" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection)
        Return dtResult
    End Function

    Friend Function GetProgramPropertyValue(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal programTypeId As Integer,
                                            ByVal propertyId As Integer, ByVal keyId As Integer) As DataTable
        Return dbUtil.List("select * from ProgramPropertyValue where ProgramTypeID=" & programTypeId & " AND PropertyID=" & propertyId & " AND KeyID=" & keyId, connection)
    End Function

    Friend Function InsertMaxDocumentNumber(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentShopId As Integer,
                                                   ByVal documentTypeId As Integer, ByVal docYear As Integer, ByVal docMonth As Integer, ByVal maxDocumentNumber As Integer) As Integer
        Dim strSQL As String
        strSQL = "Insert INTO MaxDocumentNumber(ShopID, DocType, IsDocumentOrBatch, DocumentYear, DocumentMonth, MaxDocumentNumber) " &
                 "VALUES(" & documentShopId & ", " & documentTypeId & ", 0, " & docYear & "," & docMonth & "," & maxDocumentNumber & ")"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function InsertDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                                ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal materialID As Integer, ByVal materialAmount As Decimal,
                                                ByVal percentDiscount As Decimal, ByVal amountDiscount As Decimal, ByVal pricePerUnit As Decimal, ByVal materialTax As Decimal,
                                                ByVal taxType As Integer, ByVal materialUnitSmallID As Integer, ByVal selectUnitID As Integer, ByVal selectUnitName As String,
                                                ByVal unitSmallAmount As Decimal, ByVal materialNetPrice As Decimal, ByVal materialCode As String,
                                                ByVal materialName As String, ByVal supplierMaterialCode As String, ByVal supplierMaterialName As String) As String
        Dim strSQL As String = ""
        strSQL = "Insert INTO DocDetail(DocDetailID, DocumentID, ShopID, ProductID, ProductUnit, ProductAmount, " &
                 "ProductDiscount, ProductDiscountAmount, ProductPricePerUnit,ProductTax, ProductTaxType, UnitID, " &
                 "UnitName, UnitSmallAmount, ProductNetPrice,ProductCode,ProductName,SupplierMaterialCode,SupplierMaterialName) " &
                 "Values(" & docDetailID & ", " & documentID & ", " & documentShopID & ", " & materialID & ", " & materialUnitSmallID & ", " & materialAmount &
                 "," & percentDiscount & ", " & amountDiscount & "," & pricePerUnit & ", " & materialTax & ", " & taxType & ", " & selectUnitID &
                 ",'" & selectUnitName & "', " & unitSmallAmount & "," & materialNetPrice & ",'" & materialCode & "','" & materialName & "','" & supplierMaterialCode & "','" & supplierMaterialName & "')"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function InsertDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                                ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal materialID As Integer, ByVal materialAmount As Decimal,
                                                ByVal materialUnitSmallID As Integer, ByVal selectUnitID As Integer, ByVal selectUnitName As String,
                                                ByVal unitSmallAmount As Decimal, ByVal materialCode As String, ByVal materialName As String,
                                                ByVal supplierMaterialCode As String, ByVal supplierMaterialName As String) As String
        Dim strSQL As String = ""
        strSQL = "Insert INTO DocDetail(DocDetailID, DocumentID, ShopID, ProductID, ProductUnit, ProductAmount, " &
                 "UnitID, UnitName, UnitSmallAmount, ProductCode,ProductName,SupplierMaterialCode,SupplierMaterialName) " &
                 "Values(" & docDetailID & ", " & documentID & ", " & documentShopID & ", " & materialID & ", " & materialUnitSmallID & ", " & materialAmount &
                 "," & selectUnitID & ",'" & selectUnitName & "', " & unitSmallAmount & ",'" & materialCode & "','" & materialName & "','" & supplierMaterialCode & "','" & supplierMaterialName & "')"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function InsertDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                                ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal materialID As Integer, ByVal materialAmount As Decimal,
                                                ByVal materialUnitSmallID As Integer, ByVal selectUnitID As Integer, ByVal selectUnitName As String,
                                                ByVal unitSmallAmount As Decimal, ByVal materialCode As String, ByVal materialName As String,
                                                ByVal supplierMaterialCode As String, ByVal supplierMaterialName As String,
                                                ByVal stockAmount As Decimal, ByVal diffStockAmount As Decimal, ByVal isAddRedueStock As Integer) As String
        Dim strSQL As String = ""
        strSQL = "Insert INTO DocDetail(DocDetailID, DocumentID, ShopID, ProductID, ProductUnit, ProductAmount, " &
                 "UnitID, UnitName, UnitSmallAmount, ProductCode,ProductName,SupplierMaterialCode,SupplierMaterialName,StockAmount,DiffStockAmount,IsAddRedueStock) " &
                 "Values(" & docDetailID & ", " & documentID & ", " & documentShopID & ", " & materialID & ", " & materialUnitSmallID & ", " & materialAmount &
                 "," & selectUnitID & ",'" & selectUnitName & "', " & unitSmallAmount & ",'" & materialCode & "','" & materialName & "','" & supplierMaterialCode & "','" & supplierMaterialName & "'" &
                 "," & stockAmount & "," & diffStockAmount & "," & isAddRedueStock & " )"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function InsertDocumentHeader(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                               ByVal documentShopId As Integer, ByVal documentTypeId As Integer, ByVal docMonth As Integer, ByVal docYear As Integer,
                                               ByVal newDocNumber As Integer,
                                               ByVal langID As Integer) As Integer
        Dim strSQL As String
        strSQL = "Update Document " & _
                 "Set DocumentYear = " & docYear & ", DocumentMonth = " & docMonth & ", DocumentNumber = " & newDocNumber & " " & _
                 "Where DocumentID = " & documentId & " AND ShopID = " & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function InsertAddRedueDocDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                            ByVal documentShopId As Integer, ByVal countDocumentId As Integer, ByVal countDocumentShopId As Integer, ByVal isAddRedueStock As Integer) As Integer
        Dim strSQL As String
        Dim dtResult As New DataTable
        strSQL = "select DocDetailID," & documentId & " As documentid," & documentShopId & " As ShopId,ProductID,ProductCode,ProductName,SupplierMaterialCode,SupplierMaterialName,UnitId,ProductUnit,UnitName,DiffStockAmount,DiffStockAmount" & _
                 " from docdetail where documentid =" & countDocumentId & " And shopid = " & countDocumentShopId & " And IsAddRedueStock=" & isAddRedueStock
        dtResult = dbUtil.List(strSQL, connection, objTrans)
        Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection, SqlBulkCopyOptions.Default, objTrans)
            bulkCopy.DestinationTableName = "DocDetail"
            bulkCopy.WriteToServer(dtResult)
        End Using
    End Function

    Friend Function InsertMaxDocumentId(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal docShopID As Integer,
                                               ByVal newDocId As Integer) As Integer
        Dim strSQL As String
        strSQL = "Insert INTO MaxDocumentID(MaxDocumentID, ShopID, IsDocumentOrBatch) " &
                 "VALUES(" & newDocId & ", " & docShopID & ", 0)"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function InsertMaterialDefaultPrice(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal docShopID As Integer,
                                                      ByVal vendorId As Integer, ByVal unitLargeId As Integer, ByVal defaultPrice As Decimal, ByVal unitSmallAmount As Decimal,
                                                      ByVal unitSmallId As Integer, ByVal unitSmallRatio As Integer, ByVal materialId As Integer) As Integer
        Dim strSQL As String
        strSQL = "Insert INTO MaterialDefaultPrice(MaterialID, SelectUnitLargeID, DefaultPrice, UnitSmallAmount, UnitSmallID, " &
                         "UnitSmallRatio, InventoryID, VendorID) " &
                         "VALUES(" & materialId & ", " & unitLargeId &
                         ", " & defaultPrice & ", " & unitSmallAmount & ", " & unitSmallId &
                         ", " & unitSmallRatio & ", " & docShopID & ", " & vendorId & ")"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function InsertLog(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal moduleName As String, ByVal methodName As String, ByVal errCode As String, ByVal errMsg As String) As Integer
        Dim strSQL As String = ""
        strSQL = "Insert Into pRoMiSeLog(ModuleName,MethodName,ErrorCode,ErrorMsg,LogDateTime)Values('" & moduleName & "','" & methodName & "','" & errCode & "','" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(errMsg) & "'," & FormatDateTime(Date.Now) & ")"
        Return dbUtil.sqlExecute(strSQL, connection)
    End Function

    Friend Function InsertLog(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal moduleName As String, ByVal methodName As String, ByVal errCode As String, ByVal errMsg As String) As Integer
        Dim strSQL As String = ""
        strSQL = "Insert Into pRoMiSeLog(ModuleName,MethodName,ErrorCode,ErrorMsg,LogDateTime)Values('" & moduleName & "','" & methodName & "','" & errCode & "','" & Utilitys.Utilitys.ReplaceSuitableStringForSQL(errMsg) & "'," & FormatDateTime(Date.Now) & ")"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function


    Friend Function ReSetDocumentStatus(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentId As Integer, ByVal documentShopId As Integer, ByVal documentStatus As Integer) As Integer
        Dim strSQL As String = ""
        strSQL = "Update Document Set DocumentStatus=" & documentStatus & " Where DocumentId=" & documentId & " And ShopId=" & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection)
    End Function

    Friend Function ListShiftData(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal inventoryId As Integer) As DataTable
        Dim strSQL As String
        Dim currentDate As String
        currentDate = FormatDate(Date.Now)
        strSQL = "Select DAY_ID,PERIOD_ID,SHIFT_NO From TBPERIODS Where BUS_DATE = " & currentDate & " and SHIFT_NO<>0 And POS_ID=1 ORDER BY SHIFT_NO"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function ListShiftData(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal inventoryId As Integer, ByVal receiveDate As Date) As DataTable
        Dim strSQL As String
        Dim currentDate As String
        currentDate = FormatDate(receiveDate)
        strSQL = "Select DAY_ID,PERIOD_ID,SHIFT_NO From TBPERIODS Where BUS_DATE = " & currentDate & " and SHIFT_NO<>0 And POS_ID=1 ORDER BY SHIFT_NO"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function ListShiftData(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal inventoryId As Integer, ByVal receiveDate As Date, ByVal shiftId As Integer) As DataTable
        Dim strSQL As String
        Dim currentDate As String
        currentDate = FormatDate(receiveDate)
        strSQL = "Select DAY_ID,PERIOD_ID,SHIFT_NO From TBPERIODS Where BUS_DATE = " & currentDate & "And PERIOD_ID=" & shiftId & " and SHIFT_NO<>0 And POS_ID=1 ORDER BY SHIFT_NO"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function ListPlant(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection) As DataTable
        Dim strSQL As String
        strSQL = "select * from LKDEPOT"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function ListCustomer(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection) As DataTable
        Dim strSQL As String
        strSQL = "select * from APP_DATA"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function ListBusinessPlace(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection) As DataTable
        Dim strSQL As String
        strSQL = "select * from LKBUSSINESS_PLACE"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function RefreshMonthStockForAddRedueDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer, ByVal documentShopId As Integer) As Integer
        Dim strSQL As String = ""
        strSQL = "update DocDetail set DocDetail.IsAddRedueStock=aa.IsAddRedueStock,DocDetail.StockAmount=aa.StockAmount " & _
                "from (select DocdetailId,DocumentID,ShopID, ProductId,ProductAmount,ABS(StockAmount)As StockAmount, " & _
                "case when StockAmount < 0 then 1 else 0 end As IsAddRedueStock " & _
                "from DocDetail Where documentid=" & documentId & " And ShopID=" & documentShopId & ") As aa " & _
                "where DocDetail.DocdetailId = aa.DocdetailId And DocDetail.DocumentId = aa.DocumentId And DocDetail.ShopId = aa.ShopId"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function SearchDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentTypeID As String,
                                          ByVal searchDateFrom As String, ByVal searchDateTo As String, ByVal statusDocument As Integer,
                                          ByVal searchInventoryID As String, ByVal vendorID As Integer, ByVal vendorGroupID As Integer,
                                          ByVal langID As Integer) As DataTable

        Dim strSQL As String
        Dim strProductLevel As String
        Dim strVendorGroup As String
        Dim strVendor As String
        Dim strDocumentDate As String
        Dim strStatusDocument As String

        strDocumentDate = ""
        If searchDateFrom <> "" Then
            strDocumentDate = " AND d.DocumentDate >= " & searchDateFrom
        End If
        If searchDateTo <> "" Then
            strDocumentDate &= " AND d.DocumentDate <= " & searchDateTo
        End If
        If vendorID = -1 Then
            strVendor = " "
        Else
            strVendor = " AND v.VendorID = " & vendorID & " "
        End If
        If vendorGroupID = -1 Then
            strVendorGroup = " "
        Else
            strVendorGroup = " AND v.VendorGroupID = " & vendorGroupID & " "
        End If
        If statusDocument = -1 Then
            strStatusDocument = " AND d.DocumentStatus > 0 "
        Else
            strStatusDocument = " AND d.DocumentStatus = " & statusDocument & " "
        End If
        If searchInventoryID = "" Then
            strProductLevel = " "
        Else
            strProductLevel = " AND d.ProductLevelID In (" & searchInventoryID & ") "
        End If

        strSQL = "Select d.DocumentID, d.ShopID, dt.DocumentTypeName, dt.DocumentTypeHeader, d.DocumentYear, d.DocumentMonth, d.DocumentNumber, " &
                 " d.SubTotal,d.TotalDiscount,d.TotalVAT,d.NetPrice,d.GrandTotal, " &
                 " d.DocumentDate, d.DueDate, d.Remark, d.InvoiceRef, pl.ShopName , d.ProductLevelID, d.DocumentStatus, d.ToInvID, d.FromInvID,  " &
                 " d.VendorID, d.VendorGroupID, d.VendorShopID, v.VendorCode, v.VendorName, d.DocumentIDRef, d.DocumentIDRefShopID, d.inputby,d.updateby,d.approveby, " &
                 " d.voidby,d.BusinessPlace,d.TaxInvoiceNo,d.TaxInvoiceDate " &
                 " ,dRef.DocumentStatus as RefDocumentStatus, dtRef.DocumentTypeID as RefDocumentType,dtRef.DocumentTypeName as RefDocumentTypeName" &
                 " ,dtRef.DocumentTypeHeader as RefDocumentTypeHeader, dRef.DocumentYear as RefDocumentYear, dRef.DocumentMonth as RefDocumentMonth" &
                 " ,dRef.DocumentNumber as RefDocumentNumber " &
                 " From Document d JOIN DocumentType dt ON d.DocumentTypeID = dt.DocumentTypeID AND " &
                 " dt.ComputerID = 0 AND dt.LangID = " & langID & " AND d.ShopID = dt.ShopID " &
                 " JOIN Shop_Data pl ON d.ProductLevelID = pl.ShopId " & " LEFT OUTER JOIN Document dRef ON d.DocumentIDRef = dRef.DocumentID AND d.DocumentIDRefShopID = dRef.ShopID " &
                 " LEFT OUTER JOIN DocumentType dtRef ON dRef.DocumentTypeID = dtRef.DocumentTypeID AND dRef.ShopID = dtRef.ShopID AND " &
                 " dtRef.LangID = " & langID & " AND dtRef.ComputerID = 0 " &
                 " LEFT OUTER JOIN Vendors v ON d.VendorID = v.VendorID " &
                 " Where d.DocumentTypeID In(" & documentTypeID & ") " & strProductLevel & strStatusDocument & strDocumentDate & strVendor & strVendorGroup &
                 " Order by d.DocumentDate, d.DocumentNumber, v.VendorName "
        Return dbUtil.List(strSQL, connection)

    End Function

    Friend Function SearchDocumentPTT(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentTypeID As String,
                                         ByVal searchDateFrom As String, ByVal searchDateTo As String, ByVal statusDocument As Integer,
                                         ByVal searchInventoryID As String, ByVal vendorID As Integer, ByVal vendorGroupID As Integer,
                                         ByVal langID As Integer, ByVal businessPlace As Integer, ByVal taxInvoiceNo As String,
                                         ByVal fromTaxInvoiceDate As String, ByVal toTaxInvoiceDate As String) As DataTable

        Dim strSQL As String
        Dim strProductLevel As String
        Dim strVendorGroup As String
        Dim strVendor As String
        Dim strDocumentDate As String
        Dim strStatusDocument, strInvoiceDate, strInvoiceNo, strBusinessPlace As String

        strDocumentDate = ""
        If searchDateFrom <> "" Then
            strDocumentDate = " AND d.DocumentDate >= " & searchDateFrom
        End If
        If searchDateTo <> "" Then
            strDocumentDate &= " AND d.DocumentDate <= " & searchDateTo
        End If
        strInvoiceDate = ""
        If fromTaxInvoiceDate <> "" Then
            strInvoiceDate = " AND d.TaxInvoiceDate >= " & fromTaxInvoiceDate
        End If
        If toTaxInvoiceDate <> "" Then
            strInvoiceDate &= " AND d.TaxInvoiceDate <= " & toTaxInvoiceDate
        End If
        strInvoiceNo = ""
        If taxInvoiceNo <> "" Then
            strInvoiceNo = " AND d.TaxInvoiceNo Like '%" & taxInvoiceNo & "%'"
        End If
        strBusinessPlace = ""
        If businessPlace > 0 Then
            strBusinessPlace = " AND d.BusinessPlace=" & businessPlace
        End If
        If vendorID = -1 Then
            strVendor = " "
        Else
            strVendor = " AND v.VendorID = " & vendorID & " "
        End If
        If vendorGroupID = -1 Then
            strVendorGroup = " "
        Else
            strVendorGroup = " AND v.VendorGroupID = " & vendorGroupID & " "
        End If
        If statusDocument = -1 Then
            strStatusDocument = " AND d.DocumentStatus > 0 "
        Else
            strStatusDocument = " AND d.DocumentStatus = " & statusDocument & " "
        End If
        If searchInventoryID = "" Then
            strProductLevel = " "
        Else
            strProductLevel = " AND d.ProductLevelID In (" & searchInventoryID & ") "
        End If

        strSQL = "Select d.DocumentID, d.ShopID, dt.DocumentTypeName, dt.DocumentTypeHeader, d.DocumentYear, d.DocumentMonth, d.DocumentNumber, " &
                 " d.SubTotal,d.TotalDiscount,d.TotalVAT,d.NetPrice,d.GrandTotal, " &
                 " d.DocumentDate, d.DueDate, d.Remark, d.InvoiceRef, pl.ShopName , d.ProductLevelID, d.DocumentStatus, d.ToInvID, d.FromInvID,  " &
                 " d.VendorID, d.VendorGroupID, d.VendorShopID, v.VendorCode, v.VendorName, d.DocumentIDRef, d.DocumentIDRefShopID, d.inputby, " &
                 " d.updateby,d.approveby,d.voidby,d.BusinessPlace,d.TaxInvoiceNo,d.TaxInvoiceDate " &
                 " ,dRef.DocumentStatus as RefDocumentStatus, dtRef.DocumentTypeID as RefDocumentType,dtRef.DocumentTypeName as RefDocumentTypeName" &
                 " ,dtRef.DocumentTypeHeader as RefDocumentTypeHeader, dRef.DocumentYear as RefDocumentYear, dRef.DocumentMonth as RefDocumentMonth" &
                 " ,dRef.DocumentNumber as RefDocumentNumber " &
                 " From Document d JOIN DocumentType dt ON d.DocumentTypeID = dt.DocumentTypeID AND " &
                 " dt.ComputerID = 0 AND dt.LangID = " & langID & " AND d.ShopID = dt.ShopID " &
                 " JOIN Shop_Data pl ON d.ProductLevelID = pl.ShopId " & " LEFT OUTER JOIN Document dRef ON d.DocumentIDRef = dRef.DocumentID AND d.DocumentIDRefShopID = dRef.ShopID " &
                 " LEFT OUTER JOIN DocumentType dtRef ON dRef.DocumentTypeID = dtRef.DocumentTypeID AND dRef.ShopID = dtRef.ShopID AND " &
                 " dtRef.LangID = " & langID & " AND dtRef.ComputerID = 0 " &
                 " LEFT OUTER JOIN Vendors v ON d.VendorID = v.VendorID " &
                 " Where d.DocumentTypeID In(" & documentTypeID & ") " & strProductLevel & strStatusDocument & strDocumentDate & strVendor & strVendorGroup & strInvoiceDate & strInvoiceNo & strBusinessPlace &
                 " Order by d.DocumentDate, d.DocumentNumber, v.VendorName "
        Return dbUtil.List(strSQL, connection)

    End Function


    Friend Function SearchDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentTypeID As Integer,
                                          ByVal searchDateFrom As String, ByVal searchDateTo As String, ByVal statusDocument As Integer,
                                          ByVal searchInventoryID As String, ByVal searchToInventoryId As String, ByVal langID As Integer) As DataTable

        Dim strSQL As String
        Dim strProductLevel As String
        Dim strToInventory As String
        Dim strDocumentDate As String
        Dim strStatusDocument As String

        strDocumentDate = ""
        If searchDateFrom <> "" Then
            strDocumentDate = " AND d.DocumentDate >= " & searchDateFrom
        End If
        If searchDateTo <> "" Then
            strDocumentDate &= " AND d.DocumentDate <= " & searchDateTo
        End If

        If statusDocument = -1 Then
            strStatusDocument = " AND d.DocumentStatus > 0 "
        Else
            strStatusDocument = " AND d.DocumentStatus = " & statusDocument & " "
        End If
        If searchInventoryID = "" Then
            strProductLevel = " "
        Else
            strProductLevel = " AND d.ProductLevelID In (" & searchInventoryID & ") "
        End If
        Select Case documentTypeID
            Case Is = GlobalVariable.DOCUMENTTYPE_TRANSFER
                If searchToInventoryId = -1 Then
                    strToInventory = " "
                Else
                    strToInventory = " AND d.ToInvID IN (" & searchToInventoryId & ") "
                End If
            Case Is = GlobalVariable.DOCUMENTTYPE_ROTRANSFER
                If searchToInventoryId = -1 Then
                    strToInventory = " "
                Else
                    strToInventory = " AND d.FromInvID IN (" & searchToInventoryId & ") "
                End If
        End Select

        strSQL = "Select d.DocumentID, d.ShopID, dt.DocumentTypeName, dt.DocumentTypeHeader, d.DocumentYear, d.DocumentMonth, d.DocumentNumber, " &
                 " d.SubTotal,d.TotalDiscount,d.TotalVAT,d.NetPrice,d.GrandTotal, " &
                 " d.DocumentDate, d.DueDate, d.Remark, d.InvoiceRef, pl.ShopName , d.ProductLevelID, d.DocumentStatus, d.ToInvID, d.FromInvID, " &
                 " d.VendorID, d.VendorGroupID, d.VendorShopID, v.VendorCode, v.VendorName, d.DocumentIDRef, d.DocumentIDRefShopID, d.inputby,d.updateby,d.approveby,d.voidby,d.BusinessPlace,d.TaxInvoiceNo,d.TaxInvoiceDate  " &
                 " ,dRef.DocumentStatus as RefDocumentStatus, dtRef.DocumentTypeID as RefDocumentType,dtRef.DocumentTypeName as RefDocumentTypeName" &
                 " ,dtRef.DocumentTypeHeader as RefDocumentTypeHeader, dRef.DocumentYear as RefDocumentYear, dRef.DocumentMonth as RefDocumentMonth" &
                 " ,dRef.DocumentNumber as RefDocumentNumber " &
                 " From Document d JOIN DocumentType dt ON d.DocumentTypeID = dt.DocumentTypeID AND " &
                 " dt.ComputerID = 0 AND dt.LangID = " & langID & " AND d.ShopID = dt.ShopID " &
                 " JOIN Shop_Data pl ON d.ProductLevelID = pl.ShopId " & " LEFT OUTER JOIN Document dRef ON d.DocumentIDRef = dRef.DocumentID AND d.DocumentIDRefShopID = dRef.ShopID " &
                 " LEFT OUTER JOIN DocumentType dtRef ON dRef.DocumentTypeID = dtRef.DocumentTypeID AND dRef.ShopID = dtRef.ShopID AND " &
                 " dtRef.LangID = " & langID & " AND dtRef.ComputerID = 0 " &
                 " LEFT OUTER JOIN Vendors v ON d.VendorID = v.VendorID " &
                 " Where d.DocumentTypeID = " & documentTypeID & strProductLevel & strStatusDocument & strDocumentDate & strToInventory &
                 " Order by d.DocumentDate, d.DocumentNumber"
        Return dbUtil.List(strSQL, connection)

    End Function

    Friend Function SearchDocumentForCreateNewDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal documentTypeID As Integer,
                                                       ByVal searchDateFrom As String, ByVal searchDateTo As String, ByVal statusDocument As String,
                                                       ByVal searchInventoryID As String, ByVal isTransferOrder As Boolean, ByVal langID As Integer) As DataTable

        Dim strSQL As String = ""
        Dim strProductLevel As String = ""
        Dim strDocumentDate As String = ""
        Dim strStatusDocument As String = ""
        Dim strSQLAdtional As String = ""

        strDocumentDate = ""
        If searchDateFrom <> "" Then
            strDocumentDate = " AND d.DocumentDate >= " & searchDateFrom
        End If
        If searchDateTo <> "" Then
            strDocumentDate &= " AND d.DocumentDate <= " & searchDateTo
        End If
        strStatusDocument = " AND d.DocumentStatus In(" & statusDocument & ") "
        If isTransferOrder = False Then
            strProductLevel = " AND d.ProductLevelID In (" & searchInventoryID & ") "
        Else
            strSQLAdtional = " AND d.ToInvID = " & searchInventoryID & " AND d.receiveby = 0"
        End If

        strSQL = "Select d.DocumentID, d.ShopID, dt.DocumentTypeName, dt.DocumentTypeHeader, d.DocumentYear, d.DocumentMonth, d.DocumentNumber, " &
                 " d.SubTotal,d.TotalDiscount,d.TotalVAT,d.NetPrice,d.GrandTotal, " &
                 " d.DocumentDate, d.DueDate, d.Remark, d.InvoiceRef,d.ToInvID,d.receiveby, pl.ShopName , d.ProductLevelID, d.DocumentStatus, d.ToInvID, d.FromInvID,  " &
                 " d.VendorID, d.VendorGroupID, d.VendorShopID, v.VendorCode, v.VendorName, d.DocumentIDRef, d.DocumentIDRefShopID, d.inputby,d.updateby,d.approveby,d.voidby " &
                 " ,dRef.DocumentStatus as RefDocumentStatus, dtRef.DocumentTypeID as RefDocumentType,dtRef.DocumentTypeName as RefDocumentTypeName" &
                 " ,dtRef.DocumentTypeHeader as RefDocumentTypeHeader, dRef.DocumentYear as RefDocumentYear, dRef.DocumentMonth as RefDocumentMonth" &
                 " ,dRef.DocumentNumber as RefDocumentNumber " &
                 " From Document d JOIN DocumentType dt ON d.DocumentTypeID = dt.DocumentTypeID AND " &
                 " dt.ComputerID = 0 AND dt.LangID = " & langID & " AND d.ShopID = dt.ShopID " &
                 " JOIN Shop_data pl ON d.ProductLevelID = pl.ShopId " & " LEFT OUTER JOIN Document dRef ON d.DocumentIDRef = dRef.DocumentID AND d.DocumentIDRefShopID = dRef.ShopID " &
                 " LEFT OUTER JOIN DocumentType dtRef ON dRef.DocumentTypeID = dtRef.DocumentTypeID AND dRef.ShopID = dtRef.ShopID AND " &
                 " dtRef.LangID = " & langID & " AND dtRef.ComputerID = 0 " &
                 " LEFT OUTER JOIN Vendors v ON d.VendorID = v.VendorID " &
                 " Where d.DocumentTypeID = " & documentTypeID & strProductLevel & strStatusDocument & strDocumentDate & strSQLAdtional &
                 " Order by d.DocumentDate, d.DocumentNumber, v.VendorName "
        Return dbUtil.List(strSQL, connection)

    End Function

    Friend Function SearchStatusDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection) As DataTable
        Dim strSQL As String = ""
        strSQL = "select * from DocumentStatus"
        Return dbUtil.List(strSQL, connection)
    End Function

    Public Function CreateDummyStockCard(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal inventoryid As Integer, ByVal startDate As Date, ByVal endDate As Date) As Integer

        Dim strSQL As String = ""
        Dim additionalQuery As String = ""
        Dim fristDay As Integer
        Dim fristDate As Date
        Dim strStartDate, strEndDate As String
        fristDay = startDate.Day
        
        Select Case fristDay
            Case Is = 1
                strStartDate = FormatDate(startDate)
                strEndDate = FormatDate(endDate)
                strSQL = "(select 0 AS Ordering, 0 AS DocumentTypeGroupID,'Beginning' AS GroupHeader,ProductID As MaterialID,UnitSmallAmount As NetSmallAmount," &
                        " UnitSmallAmount AS TotalAmount,ProductNetPrice AS ProductNetPrice, 0 As CalculateStandardProfitLoss FROM document aa, docdetail bb, documenttype dt " &
                        " where aa.DocumentID=bb.DocumentID AND aa.ShopID=bb.ShopID AND aa.DocumentTypeID=10  AND aa.DocumentStatus=2 AND aa.DocumentTypeID=dt.DocumentTypeID " &
                        " AND dt.LangID=1 AND dt.ShopID=aa.ProductLevelID AND aa.ProductLevelID=" + inventoryid.ToString + " AND aa.DocumentDate = " + strStartDate + ") "
            Case Else
                fristDate = startDate.AddDays(1 - startDate.Day)
                strStartDate = FormatDate(fristDate)
                strEndDate = FormatDate(startDate)
                strSQL = "(select 0 AS Ordering, 0 AS DocumentTypeGroupID,'Beginning' AS GroupHeader,ProductID As MaterialID," &
                     " sum(UnitSmallAmount * dt.movementinstock) As NetSmallAmount, sum(UnitSmallAmount * dt.movementinstock) AS TotalAmount,sum(bb.ProductNetPrice * dt.movementinstock) as ProductNetPrice, 0 as CalculateStandardProfitLoss" &
                     " FROM document aa, docdetail bb, documenttype dt " &
                     " where aa.DocumentID=bb.DocumentID AND aa.ShopID=bb.ShopID AND aa.DocumentStatus=2 AND aa.DocumentTypeID=dt.DocumentTypeID " &
                     " AND dt.LangID=1 AND dt.ShopID=aa.ProductLevelID AND aa.ProductLevelID=" + inventoryid.ToString + " AND aa.DocumentDate >= " + strStartDate + " AND aa.DocumentDate < " + strEndDate + " " &
                     " group by productid )"
        End Select
        strStartDate = FormatDate(startDate)
        strEndDate = FormatDate(endDate)

        strSQL &= " UNION "
        strSQL &= "(select d.Ordering, d.DocumentTypeGroupID,d.GroupHeader, ProductID As MaterialID ,sum(e.MovementInStock*b.UnitSmallAmount) As NetSmallAmount, " &
                    " sum(b.UnitSmallAmount) AS TotalAmount, sum(ProductNetPrice) As ProductNetPrice,e.CalculateStandardProfitLoss from document a, docdetail b, documentTypeGroupValue c, " &
                    " DocumentTypeGroup d, DocumentType e where a.DocumentID=b.DocumentID AND a.ShopID=b.ShopID AND a.DocumentTypeID=c.DocumentTypeID  " &
                    " AND c.DocumentTypeGroupID = d.DocumentTypeGroupID  AND a.DocumentTypeID=e.DocumentTypeID AND a.ShopID=e.ShopID  AND d.Ordering > 0 AND a.DocumentStatus=2 " &
                    " AND a.ProductLevelID=" + inventoryid.ToString + " AND e.LangID=1  AND a.DocumentDate >= " + strStartDate + " AND a.DocumentDate <= " + strEndDate +
                    " group by d.Ordering,d.DocumentTypeGroupID,d.GroupHeader, ProductID,e.CalculateStandardProfitLoss)"
        strSQL &= " UNION "
        strSQL &= "(select d.Ordering,d.DocumentTypeGroupID,d.GroupHeader, ProductID As MaterialID ,sum(e.MovementInStock*b.UnitSmallAmount) As NetSmallAmount, " &
                    " sum(b.UnitSmallAmount) AS TotalAmount, sum(ProductNetPrice) As ProductNetPrice,e.CalculateStandardProfitLoss from document a, docdetail b, documentTypeGroupValue " &
                    " c, DocumentTypeGroup d, DocumentType e where a.DocumentID=b.DocumentID AND a.ShopID=b.ShopID AND a.DocumentTypeID=c.DocumentTypeID " &
                    " AND c.DocumentTypeGroupID = d.DocumentTypeGroupID  AND a.DocumentTypeID=e.DocumentTypeID AND a.ShopID=e.ShopID  AND d.Ordering < 0 " &
                    " AND a.DocumentStatus=2 AND a.ProductLevelID=" + inventoryid.ToString + " AND e.LangID=1  " &
                    " AND a.DocumentDate >= " + strStartDate + " AND a.DocumentDate <= " + strEndDate + " " &
                    " group by d.Ordering,d.DocumentTypeGroupID,d.GroupHeader, ProductID,CalculateStandardProfitLoss)"

        dbUtil.sqlExecute("IF OBJECT_ID('DummyStockCard', 'U') IS NOT NULL DROP TABLE DummyStockCard", connection)
        dbUtil.sqlExecute("create table DummyStockCard (Ordering int, DocumentTypeGroupID int, GroupHeader varchar(50), MaterialID int, NetSmallAmount decimal(18,4),TotalAmount decimal(18,4), ProductNetPrice decimal(18,4),CalculateStandardProfitLoss int)", connection)
        'Check Log
        'DocumentSQL.InsertLog(dbUtil, connection, "StockCard", "SearchByCode", "77", strSQL.ToString)
        Return dbUtil.sqlExecute("insert into DummyStockCard(Ordering,DocumentTypeGroupID,GroupHeader,MaterialID,NetSmallAmount,TotalAmount,ProductNetPrice,CalculateStandardProfitLoss) " + strSQL, connection)
       
    End Function

    Friend Function UpdateDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                          ByVal documentShopId As Integer, ByVal productLevelID As Integer, ByVal vendorID As Integer, ByVal vendorGroupID As Integer,
                                          ByVal docDate As String, ByVal remark As String, ByVal invoiceRef As String, ByVal termOfPayment As Integer,
                                          ByVal creditDay As Integer, ByVal dueDate As String, ByVal VATPercent As Decimal, ByVal updateDate As String,
                                          ByVal updateBy As Integer) As Integer
        Dim strSQL As String
        strSQL = "Update Document " & _
                 "Set ProductLevelID = " & productLevelID & ", ToInvID = " & productLevelID & ", VendorID = " & vendorID & _
                 ", VendorGroupID = " & vendorGroupID & ", VendorShopID = 0" & _
                 ", DocumentDate = " & docDate & ", Remark = '" & remark & "', InvoiceRef = '" & invoiceRef & _
                 "', TermOfPayment = " & termOfPayment & ", CreditDay = " & creditDay & ", DueDate = " & dueDate & _
                 ", VATPercent = " & VATPercent & ", UpdateBy = " & updateBy & ", UpdateDate = " & updateDate & _
                 ", AlreadyExportToHQ = 0, AlreadyExportToBranch = 0 " & _
                 "Where DocumentID = " & documentId & " AND ShopID = " & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocumentDeliveryCost(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                         ByVal documentShopId As Integer, ByVal deliveryTaxClass As Integer, ByVal deliveryCost As Decimal, ByVal deliveryTax As Decimal,
                                         ByVal deliveryNetPrice As Decimal) As Integer
        Dim strSQL As String
        strSQL = "Update Document " & _
                 "Set TransferTotal= " & deliveryCost & ",TransferTaxClass=" & deliveryTaxClass & ",TransferVAT=" & deliveryTax & ",TransferNetPrice=" & deliveryNetPrice &
                 "Where DocumentID = " & documentId & " AND ShopID = " & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                          ByVal documentShopId As Integer, ByVal toInventoryId As Integer, ByVal docDate As String, ByVal dueDate As String,
                                          ByVal remark As String, ByVal invoiceRef As String, ByVal updateDate As String, ByVal updateBy As Integer) As Integer
        Dim strSQL As String
        strSQL = "Update Document " & _
                 "Set ProductLevelID = " & documentShopId & ", ToInvID = " & toInventoryId &
                 ", DocumentDate = " & docDate & ", Remark = '" & remark & "', InvoiceRef = '" & invoiceRef & _
                 "', DueDate = " & dueDate & ", UpdateBy = " & updateBy & ", UpdateDate = " & updateDate & _
                 ", AlreadyExportToHQ = 0, AlreadyExportToBranch = 0 " & _
                 "Where DocumentID = " & documentId & " AND ShopID = " & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateMaxDocumentNumber(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentShopId As Integer,
                                                    ByVal documentTypeId As Integer, ByVal docYear As Integer, ByVal docMonth As Integer, ByVal maxDocumentNumber As Integer) As Integer
        Dim strSQL As String
        strSQL = "Update MaxDocumentNumber Set MaxDocumentNumber = " & maxDocumentNumber & " " & _
                       "Where ShopID = " & documentShopId & " AND DocType = " & documentTypeId & " AND " & _
                       " DocumentYear = " & docYear & " AND DocumentMonth = " & docMonth & " AND IsDocumentOrBatch = 0 "
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocumentStatus(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                                ByVal documentShopId As Integer, ByVal docStatus As Integer, ByVal updateDate As String, ByVal updateBy As Integer) As Integer
        Dim strSQL As String
        strSQL = "Update Document Set NewSend=1, DocumentStatus = " & docStatus & ", UpdateBy = " & updateBy & _
                 ", UpdateDate = " & updateDate & ", AlreadyExportToHQ = 0, AlreadyExportToBranch = 0 " & _
                 "Where DocumentID = " & documentId & " AND ShopID = " & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocumentStatus(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                         ByVal documentShopId As Integer, ByVal docStatus As Integer, ByVal updateDate As String, ByVal updateBy As Integer,
                                         ByVal receiveBy As Integer) As Integer
        Dim strSQL As String
        strSQL = "Update Document Set NewSend=1, DocumentStatus = " & docStatus & ", UpdateBy = " & updateBy & ", ReceiveBy = " & receiveBy & _
                 ", UpdateDate = " & updateDate & ", AlreadyExportToHQ = 0, AlreadyExportToBranch = 0 " & _
                 "Where DocumentID = " & documentId & " AND ShopID = " & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateReferenceNetPriceInDocDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction,
                                                              ByVal documentId As Integer, ByVal documentShopId As Integer, ByVal docDetailID As Integer,
                                                              ByVal refNetPrice As Decimal, ByVal refProductTax As Decimal) As Integer
        Dim strSQL As String
        strSQL = "Update DocDetail " &
                 "Set ReferenceNetPrice = " & refNetPrice & ", ReferenceProductTax = " & refProductTax & " " &
                 "Where DocDetailID = " & docDetailID & " AND DocumentID = " & documentId & " AND ShopID = " & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocSummaryIntoDocument(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction,
                                                              ByVal documentId As Integer, ByVal documentShopId As Integer)
        Dim strSQL As String = ""
        strSQL = "Update Document  Set Document.SubTotal=dd.SubTotal,Document.TotalVAT=dd.TotalVAT," &
                 " Document.NetPrice=dd.NetPrice,Document.TotalDiscount=dd.TotalDiscount,Document.GrandTotal=dd.GrandTotal" &
                 " From (Select DocumentID,ShopID,Round(Sum(ProductAmount * ProductPricePerUnit),2) As SubTotal" &
                    ",Round(Sum((((ProductAmount * ProductPricePerUnit) * ProductDiscount)/100)+ ProductDiscountAmount),2) As TotalDiscount" &
                    ",Round(Sum(ProductTax),2) As TotalVAT  " &
                    ",Round(Sum(ProductNetPrice),2) As NetPrice " &
                    ",Round(Sum(ProductTax + ProductNetPrice),2) As GrandTotal    " &
                    " From DocDetail Where DocumentId=" & documentId & " And ShopID=" & documentShopId &
                    " Group by DocumentID,ShopID) dd  " &
                    " Where Document.DocumentID = dd.DocumentID And Document.ShopID = dd.ShopID"
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                                ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal materialID As Integer, ByVal materialAmount As Decimal,
                                                ByVal percentDiscount As Decimal, ByVal amountDiscount As Decimal, ByVal pricePerUnit As Decimal, ByVal materialTax As Decimal,
                                                ByVal taxType As Integer, ByVal materialUnitSmallID As Integer, ByVal selectUnitID As Integer, ByVal selectUnitName As String,
                                                ByVal unitSmallAmount As Decimal, ByVal materialNetPrice As Decimal, ByVal materialCode As String,
                                                ByVal materialName As String, ByVal supplierMaterialCode As String, ByVal supplierMaterialName As String) As Integer
        Dim strSQL As String
        strSQL = "Update DocDetail " &
                 "Set ProductID = " & materialID & ", ProductUnit = " & materialUnitSmallID &
                 ", ProductAmount = " & materialAmount & ", ProductDiscount = " & percentDiscount &
                 ", ProductDiscountAmount = " & amountDiscount & ", ProductPricePerUnit = " & pricePerUnit &
                 ", ProductTax = " & materialTax & ", ProductTaxType = " & taxType &
                 ", UnitID = " & selectUnitID & ", UnitName = '" & selectUnitName &
                 "',UnitSmallAmount = " & unitSmallAmount & ", ProductNetPrice = " & materialNetPrice &
                 ", ProductCode='" & materialCode & "',ProductName='" & materialName & "',SupplierMaterialCode='" & supplierMaterialCode & _
                 "',SupplierMaterialName='" & supplierMaterialName & "'" & _
                 " Where DocDetailID = " & docDetailID & " AND DocumentID = " & documentID & " AND ShopID = " & documentShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                                ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal materialID As Integer, ByVal materialAmount As Decimal,
                                                ByVal materialUnitSmallID As Integer, ByVal selectUnitID As Integer, ByVal selectUnitName As String,
                                                ByVal unitSmallAmount As Decimal, ByVal materialCode As String, ByVal materialName As String,
                                                ByVal supplierMaterialCode As String, ByVal supplierMaterialName As String) As Integer
        Dim strSQL As String
        strSQL = "Update DocDetail " &
                 "Set ProductID = " & materialID & ", ProductUnit = " & materialUnitSmallID &
                 ", ProductAmount = " & materialAmount &
                 ", UnitID = " & selectUnitID & ", UnitName = '" & selectUnitName &
                 "',UnitSmallAmount = " & unitSmallAmount & ", ProductCode='" & materialCode &
                 "',ProductName='" & materialName & "',SupplierMaterialCode='" & supplierMaterialCode & _
                 "',SupplierMaterialName='" & supplierMaterialName & "'" & _
                 " Where DocDetailID = " & docDetailID & " AND DocumentID = " & documentID & " AND ShopID = " & documentShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateDocumentDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentID As Integer,
                                               ByVal documentShopID As Integer, ByVal docDetailID As Integer, ByVal materialID As Integer, ByVal materialAmount As Decimal,
                                               ByVal materialUnitSmallID As Integer, ByVal selectUnitID As Integer, ByVal selectUnitName As String,
                                               ByVal unitSmallAmount As Decimal, ByVal materialCode As String, ByVal materialName As String,
                                               ByVal supplierMaterialCode As String, ByVal supplierMaterialName As String,
                                               ByVal stockAmount As Decimal, ByVal diffStockAmount As Decimal, ByVal isAddRedueStock As Integer) As Integer
        Dim strSQL As String
        strSQL = "Update DocDetail " &
                 "Set ProductID = " & materialID & ", ProductUnit = " & materialUnitSmallID &
                 ", ProductAmount = " & materialAmount &
                 ", UnitID = " & selectUnitID & ", UnitName = '" & selectUnitName &
                 "',UnitSmallAmount = " & unitSmallAmount & ", ProductCode='" & materialCode &
                 "',ProductName='" & materialName & "',SupplierMaterialCode='" & supplierMaterialCode & _
                 "',SupplierMaterialName='" & supplierMaterialName & "'" & _
                 ",StockAmount=" & stockAmount & _
                 ",DiffStockAmount=" & diffStockAmount & _
                 ",IsAddRedueStock= " & isAddRedueStock & _
                 " Where DocDetailID = " & docDetailID & " AND DocumentID = " & documentID & " AND ShopID = " & documentShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function TestReferedDocumentHasMaterialAmountLeft(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal docRefID As Integer, ByVal docRefShopID As Integer, ByVal staffID As Integer) As Boolean
        Dim strSQL As String
        Dim dtResult As DataTable
        'Get DocDetail from DocRef
        strSQL = "IF OBJECT_ID('DocDetailReferFrom" & staffID & "', 'U') IS NOT NULL DROP TABLE DocDetailReferFrom" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection, objTrans)
        strSQL = "Create Table DocDetailReferFrom" & staffID & _
                " (MaterialID int NOT NULL, MaterialSmallAmount decimal(18,4), UnitSmallID int NOT NULL); "
        dbUtil.sqlExecute(strSQL, connection, objTrans)
        'DocDetail from Refered Document (Original Document)
        strSQL = "Insert INTO DocDetailReferFrom" & staffID & " " & _
                 "Select ProductID, Sum(UnitSmallAmount), ProductUnit " & _
                 "From DocDetail " & _
                 "Where DocumentID = " & docRefID & " AND ShopID = " & docRefShopID & " " & _
                 "Group by ProductID, ProductUnit "
        dbUtil.sqlExecute(strSQL, connection, objTrans)
        'DocDetail From Document that is refer from this document
        strSQL = "Insert INTO DocDetailReferFrom" & staffID & " " & _
                 "Select dd.ProductID, - Sum(dd.UnitSmallAmount), ProductUnit " & _
                 "From Document d, DocDetail dd, DocDetailReferFrom" & staffID & " dr " & _
                 "Where d.DocumentIDRef = " & docRefID & " AND d.DocumentIDRefShopID = " & docRefShopID & " AND " & _
                 " d.DocumentID = dd.DocumentID AND dd.ShopID = d.ShopID AND d.DocumentStatus NOT IN (0,99) AND " & _
                 " dd.ProductID = dr.MaterialID AND dd.ProductUnit = dr.UnitSmallID " & _
                 "Group by dd.ProductID, dd.ProductUnit "
        dbUtil.sqlExecute(strSQL, connection, objTrans)

        'Get Only Summary Material That has Amount > 0
        strSQL = "Select MaterialID, Sum(MaterialSmallAmount), UnitSmallID " & _
                 "From DocDetailReferFrom" & staffID & " " & _
                 "Group by MaterialID, UnitSmallID " & _
                 "Having Sum(MaterialSmallAmount) > 0 "
        dtResult = dbUtil.List(strSQL, connection, objTrans)

        strSQL = "IF OBJECT_ID('DocDetailReferFrom" & staffID & "', 'U') IS NOT NULL DROP TABLE DocDetailReferFrom" & staffID & ";"
        dbUtil.sqlExecute(strSQL, connection, objTrans)

        If dtResult.Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Friend Function UpdateMaterialDefaultPrice(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal docShopID As Integer,
                                                      ByVal vendorId As Integer, ByVal unitLargeId As Integer, ByVal defaultPrice As Decimal, ByVal unitSmallAmount As Decimal,
                                                      ByVal unitSmallId As Integer, ByVal unitSmallRatio As Integer, ByVal materialId As Integer) As Integer
        Dim strSQL As String
        strSQL = "Update MaterialDefaultPrice " &
                         "Set SelectUnitLargeID = " & unitLargeId & ", DefaultPrice = " & defaultPrice &
                         ", UnitSmallAmount = " & unitSmallAmount & ", UnitSmallID = " & unitSmallId &
                         ", UnitSmallRatio = " & unitSmallRatio &
                         " Where MaterialID = " & materialId &
                         " AND InventoryID = " & docShopID &
                         " AND VendorID = " & vendorId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateMaxDocumentId(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal docShopID As Integer,
                                               ByVal newDocId As Integer) As Integer
        Dim strSQL As String
        strSQL = " Update MaxDocumentID Set MaxDocumentID = " & newDocId &
                    " Where ShopID = " & docShopID
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateMaterialRemainAmountForNewDocumentByDocumentReference(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                                                                       ByVal documentShopId As Integer, ByVal docDetailId As Integer, ByVal productAmount As Decimal, ByVal productTax As Decimal,
                                                                                       ByVal unitSmallAmount As Decimal, ByVal productNetPrice As Decimal, ByVal prepareSmallAmount As Decimal) As Integer
        Dim strSQL As String
        strSQL = "Update DocDetail Set ProductAmount = " & productAmount &
                        ", ProductTax = " & productTax & ", UnitSmallAmount = " & unitSmallAmount &
                        ", ProductNetPrice = " & productNetPrice & ", DefaultInCompare = 1, PrepareSmallAmount =" & prepareSmallAmount &
                        "Where DocumentID = " & documentId & " AND ShopID = " & documentShopId & " AND " & _
                        " DocDetailID = " & docDetailId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function GetCurrentStock(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal startDate As Date, ByVal endDate As Date, ByVal shopId As Integer, ByVal materialId As Integer) As DataTable

        Dim dtDateBeginStock As New DataTable
        Dim dtStockData As New DataTable
        Dim strSQL As String = ""
        Dim dateBegin As Date

        dtDateBeginStock = dbUtil.List("SELECT MAX(documentdate) AS maxtransferdate FROM document WHERE documenttypeid=10 AND productlevelid=" & shopId & " AND documentstatus=2", connection)
        If dtDateBeginStock.Rows.Count > 0 Then
            If Not IsDBNull(dtDateBeginStock.Rows(0)("maxtransferdate")) Then
                dateBegin = dtDateBeginStock.Rows(0)("maxtransferdate")
            Else
                dateBegin = startDate
            End If
        Else
            dateBegin = startDate
        End If
        strSQL = " SELECT dd.productid, SUM(dd.unitsmallamount * dt.movementinstock) AS qty" &
              " FROM document d, docdetail dd, documenttype dt" &
              " WHERE d.documentid=dd.documentid AND d.shopid=dd.shopid AND " &
              " d.documentdate>=" & FormatDate(dateBegin) & " AND d.documentdate<=" & FormatDate(endDate) &
              " AND d.documentstatus=2 AND d.productlevelid= " & shopId & " AND dt.shopid=d.shopid AND" &
              " dt.documenttypeid = d.documenttypeid And dt.langid = 2 "
        If materialId > 0 Then
            strSQL &= " AND dd.ProductId=" & materialId
        End If
        strSQL &= " GROUP BY dd.productid"
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function UpdateDocumentRef(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                      ByVal documentShopId As Integer, ByVal documentIdRef As Integer, ByVal documentIDRefShopID As Integer) As Integer
        Dim strSQL As String = ""
        strSQL = "Update Document Set DocumentIdRef=" & documentIdRef & ",DocumentIdRefShopID=" & documentIDRefShopID & " Where DocumentId=" & documentId & " And ShopId=" & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function UpdateStockAtDateTime(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer,
                                     ByVal documentShopId As Integer) As Integer
        Dim strSQL As String = ""
        strSQL = "Update document Set StockAtDateTime=" & FormatDateTime(Date.Now) & " Where DocumentId=" & documentId & " And ShopId=" & documentShopId
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

End Module
