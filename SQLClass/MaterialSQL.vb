Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Module MaterialSQL

    Friend Function ListMaterialGroup(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection) As DataTable
        Dim strSQL As String
        strSQL = " Select Distinct mg.MaterialGroupID, mg.MaterialGroupType, mg.MaterialGroupCode, mg.MaterialGroupName " &
                 " From MaterialGroup mg, MaterialDept md, Materials m " &
                 " Where mg.Deleted = 0 AND mg.MaterialGroupID = md.MaterialGroupID AND " &
                 "  md.MaterialDeptID = m.MaterialDeptID" &
                 " Order by mg.MaterialGroupID "
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function ListMaterialDept(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection, ByVal groupID As Integer) As DataTable
        Dim strSQL As String
        Dim strGroup As String
        strGroup = " "
        If groupID <> -1 Then
            strGroup &= " AND MaterialGroupID = " & groupID & " "
        End If
        strSQL = " Select Distinct md.MaterialDeptID, md.MaterialGroupID, md.MaterialDeptCode, md.MaterialDeptName  " &
                     " From MaterialDept md, Materials m " &
                     " Where md.Deleted = 0 " & strGroup & " AND md.MaterialDeptID = m.MaterialDeptID " &
                     " Order by md.MaterialDeptID "
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function ListMaterial(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection, ByVal groupID As Integer, ByVal deptID As Integer) As DataTable
        Dim strSQL, strDept As String
        If deptID <> 0 Then
            strDept = " AND MaterialDeptID = " & deptID & " "
        Else
            strDept = " "
        End If
        If (deptID = 0) And (groupID <> 0) Then
            strSQL = " Select m.MaterialID, m.MaterialCode, m.MaterialName, m.MaterialDeptID, m.MaterialTaxType, m.UnitSmallID " &
                    " From Materials m, MaterialDept md " &
                    " Where m.Deleted = 0  AND m.MaterialDeptID = md.MaterialDeptID AND md.MaterialGroupID = " & groupID & strDept
        Else
            strSQL = " Select MaterialID, MaterialCode, MaterialName, MaterialDeptID, MaterialTaxType, UnitSmallID " &
                              " From Materials " &
                              " Where Deleted = 0  " & strDept
        End If
        strSQL &= " Union " &
                 " Select m.MaterialID,ur.PTTCode As MaterialCode, ur.PTTName As MaterialName, m.MaterialDeptID, m.MaterialTaxType, m.UnitSmallID" &
                 " From Materials m, unitratio ur " &
                 " Where m.Deleted = 0 And m.UnitSmallID = ur.UnitSmallID " & strDept & _
                 " And (ur.PTTCode Is Not Null Or ur.PTTCode <>'')"
        strSQL &= " Order by MaterialCode, MaterialName "
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function ListMaterialUnit(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection) As DataTable
        Dim strSQL As String
        strSQL = "Select m.MaterialID, ul.UnitLargeID, ur.UnitSmallID, ul.UnitLargeName, us.UnitSmallName, ur.UnitLargeRatio, ur.UnitSmallRatio," &
                 "0 as IsDefault, 0 as IsLockPrice " &
                 "From Materials m, UnitLarge ul, UnitRatio ur, UnitSmall us " &
                 "Where ur.Deleted = 0 AND ur.UnitSmallID = m.UnitSmallID And ur.UnitLargeID = ul.UnitLargeID And " &
                 "ur.UnitSmallID = us.UnitSmallID  " &
                 "Order by m.MaterialID, ur.UnitSmallRatio"
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function ListMaterialUnit(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection, ByVal materialId As Integer) As DataTable
        Dim strSQL As String
        strSQL = "Select m.MaterialID, ul.UnitLargeID, ur.UnitSmallID, ul.UnitLargeName, us.UnitSmallName, ur.UnitLargeRatio, ur.UnitSmallRatio," &
                 "0 as IsDefault, 0 as IsLockPrice " &
                 "From Materials m, UnitLarge ul, UnitRatio ur, UnitSmall us " &
                 "Where ur.Deleted = 0 AND ur.UnitSmallID = m.UnitSmallID And ur.UnitLargeID = ul.UnitLargeID And m.MaterialID=" & materialId & " AND " &
                 "ur.UnitSmallID = us.UnitSmallID  " &
                 "Order by m.MaterialID, ur.UnitSmallRatio"
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function GetMaterialDetailFromCode(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection, ByVal materialCode As String, ByVal isSearchInUnitRatio As Boolean) As DataTable
        Dim strSQL As String
        Dim strWhereGroup As String
        strWhereGroup = " "
        If isSearchInUnitRatio = False Then
            strSQL = " Select m.MaterialID, m.MaterialCode, m.MaterialName, m.MaterialDeptID, m.MaterialTaxType, m.UnitSmallID " &
                    " From Materials m " &
                    " Where m.MaterialCode Like'%" & materialCode & "%' AND m.Deleted = 0  " & strWhereGroup
            strSQL &= " Union " &
               " Select m.MaterialID,ur.PTTCode As MaterialCode, ur.PTTName As MaterialName, m.MaterialDeptID, m.MaterialTaxType, m.UnitSmallID" &
               " From Materials m, unitratio ur " &
               " Where m.Deleted = 0 And m.UnitSmallID = ur.UnitSmallID  " & _
               " And (ur.PTTCode Is Not Null Or ur.PTTCode <>'') And ur.PTTCode Like '%" & materialCode & "%' " & strWhereGroup
        Else
            strSQL = " Select m.*, ur.UnitLargeID as SelectUnitLargeID, ur.UnitLargeRatio, ur.UnitSmallRatio " &
                    " From UnitRatio ur, Materials m " &
                    " Where ur.MaterialUnitRatioCode = '" & materialCode & "' AND ur.UnitSmallID = m.UnitSmallID AND " &
                    " ur.Deleted = 0 AND m.Deleted = 0  " & strWhereGroup
        End If
        strSQL &= " Order By m.MaterialCode, m.MaterialName "
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function GetMaterialDetail(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection, ByVal materialID As Integer) As DataTable
        Dim strSQL As String
        strSQL = " Select * " & _
                 " From Materials " & _
                 " Where MaterialID = " & materialID
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function GetMaterialDefaultPrice(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal vendorId As Integer) As DataTable
        Dim strSQL As String
        strSQL = "select * from MaterialDefaultPrice Where VendorID=" & vendorId
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function SearchMaterialByName(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection, ByVal keyWord As String) As DataTable
        Dim strSQL As String
        strSQL = " Select m.MaterialID, m.MaterialCode, m.MaterialName, m.MaterialDeptID, m.MaterialTaxType, m.UnitSmallID " &
                          " From Materials m " &
                          " Where m.MaterialName Like '%" & keyWord & "%' AND m.Deleted = 0  "
        strSQL &= " Union " &
                    " Select m.MaterialID,ur.PTTCode As MaterialCode, ur.PTTName As MaterialName, m.MaterialDeptID, m.MaterialTaxType, m.UnitSmallID" &
                    " From Materials m, unitratio ur " &
                    " Where m.Deleted = 0 And m.UnitSmallID = ur.UnitSmallID  " & _
                    " And (ur.PTTCode Is Not Null Or ur.PTTCode <>'') And ur.PTTName Like '%" & keyWord & "%' " &
                    " Order by m.MaterialName, m.MaterialCode "
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function SearchMaterialByCode(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection, ByVal keyWord As String) As DataTable
        Dim strSQL As String
        strSQL = " Select m.MaterialID, m.MaterialCode, m.MaterialName, m.MaterialDeptID, m.MaterialTaxType, m.UnitSmallID " &
                          " From Materials m " &
                          " Where m.MaterialCode Like '%" & keyWord & "%' AND m.Deleted = 0  " &
                          " Order by m.MaterialCode, m.MaterialName "
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function ListMaterialUnitLarge(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection) As DataTable
        Dim strSQL As String
        strSQL = "Select m.MaterialID, ul.UnitLargeID, ur.UnitSmallID, ul.UnitLargeName, us.UnitSmallName, ur.UnitLargeRatio, ur.UnitSmallRatio," &
                "0 as IsDefault, 0 as IsLockPrice " &
                "From Materials m, UnitLarge ul, UnitRatio ur, UnitSmall us " &
                "Where ur.Deleted = 0 AND ur.UnitSmallID = m.UnitSmallID And ur.UnitLargeID = ul.UnitLargeID And " &
                "ur.UnitSmallID = us.UnitSmallID  " &
                "Order by m.MaterialID, ur.UnitSmallRatio"
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function GetMaterialDetailAndUnitRatio(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection, ByVal materialID As Integer, ByVal unitLargeID As Integer,
                                                         ByVal isIncludeUnitSmallName As Boolean) As DataTable
        Dim strSQL As String
        If isIncludeUnitSmallName = False Then
            strSQL = "Select m.*, ul.UnitLargeName, ul.UnitLargeID as SelectUnitID, ur.UnitLargeRatio, ur.UnitSmallRatio,ur.PTTCode,ur.PTTName " & _
                    "From Materials m, UnitRatio ur, UnitLarge ul " & _
                    "Where m.MaterialID = " & materialID & " AND m.UnitSmallID = ur.UnitSmallID AND " & _
                    " ul.UnitLargeID = " & unitLargeID & " AND ur.UnitLargeID = ul.UnitLargeID  "
        Else
            strSQL = "Select m.*, us.UnitSmallName, ul.UnitLargeName, ul.UnitLargeID as SelectUnitID, ur.UnitLargeRatio, ur.UnitSmallRatio " & _
                    "From Materials m, UnitSmall us, UnitRatio ur, UnitLarge ul " & _
                    "Where m.MaterialID = " & materialID & " AND m.UnitSmallID = ur.UnitSmallID AND " & _
                    " ul.UnitLargeID = " & unitLargeID & " AND ur.UnitLargeID = ul.UnitLargeID AND us.UnitSmallID = m.UnitSmallID  "
        End If
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function GetMaterialTaxType(ByVal dbUtil As CDBUtil, ByVal objCnn As SqlConnection, ByVal taxTypeList As String) As DataTable
        Dim strSQL As String = ""
        strSQL = "select * from MaterialTaxtype where MaterialTaxtype in(" & taxTypeList & ");"
        Return dbUtil.List(strSQL, objCnn)
    End Function

    Friend Function GetMaterialFromCountStockSetting(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal tableName As String) As DataTable
        Dim strSQL As String = ""
        strSQL = "Select Distinct m.MaterialID,Case When st.SelectUnitLargeId is null Then m.UnitSmallID Else st.SelectUnitLargeId End As UnitID " &
                 "From Materials m inner join " & tableName & " d On m.MaterialId=d.MaterialId " &
                 "left join MaterialDocumentTypeUnitSetting st on d.MaterialID=st.MaterialID And st.DocumentTypeId=7 And IsDefault=1 " &
                 "Where m.Deleted = 0 "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function AutoAddDailyStockMaterial(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer, ByVal shopId As Integer)
        Dim strSQL As String = ""

        strSQL = "insert into DailyStockMaterial(ShopID,MaterialID) " & _
                 "select distinct dd.ShopID,dd.ProductID from DocDetail dd left join DailyStockMaterial m on dd.ProductID=m.MaterialID " & _
                 "where dd.DocumentId=" & documentId & " And dd.ShopID=" & shopId & " And m.MaterialID is null "
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function AutoAddMonthlyStockMaterial(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer, ByVal shopId As Integer)
        Dim strSQL As String = ""

        strSQL = "insert into MonthlyStockMaterial(ShopID,MaterialID) " & _
                 "select  distinct dd.ShopID,dd.ProductID from DocDetail dd left join DailyStockMaterial m on dd.ProductID=m.MaterialID " & _
                 "where dd.DocumentId=" & documentId & " And dd.ShopID=" & shopId & " And m.MaterialID is null "
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function

    Friend Function AutoAddWeeklyStockMaterial(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal objTrans As SqlTransaction, ByVal documentId As Integer, ByVal shopId As Integer)
        Dim strSQL As String = ""

        strSQL = "insert into WeeklyStockMaterial(ShopID,MaterialID) " & _
                 "select distinct  dd.ShopID,dd.ProductID from DocDetail dd left join DailyStockMaterial m on dd.ProductID=m.MaterialID " & _
                 "where dd.DocumentId=" & documentId & " And dd.ShopID=" & shopId & " And m.MaterialID is null "
        Return dbUtil.sqlExecute(strSQL, connection, objTrans)
    End Function


End Module
