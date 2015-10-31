Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Module InventorySQL

    Friend Function ListInventory(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal orderBy As Integer) As DataTable
        Dim strSQL As String
        strSQL = "select ShopID,ShopCode,ShopName " &
                 "from shop_data " &
                 "Where  Deleted = 0 AND IsInv=1 "
        Select Case orderBy
            Case GlobalVariable.ORDERINVENTORY_BYNAME
                strSQL &= " Order by ShopName,ShopCode "
            Case GlobalVariable.ORDERINVENTORY_BYCODEANDNAME
                strSQL &= " Order by ShopCode,ShopName "
            Case GlobalVariable.ORDERINVENTORY_BYID
                strSQL &= " Order by ShopID "
            Case Else
                strSQL &= "Order by ShopName,ShopCode"
        End Select
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetInventory(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal orderBy As Integer) As DataTable
        Dim strSQL As String
        strSQL = "select ShopID,ShopCode,ShopName " &
                 "from shop_data " &
                 "Where  Deleted = 0 AND IsInv=1 "
        Select Case orderBy
            Case GlobalVariable.ORDERINVENTORY_BYNAME
                strSQL &= " Order by ShopName,ShopCode "
            Case GlobalVariable.ORDERINVENTORY_BYCODEANDNAME
                strSQL &= " Order by ShopCode,ShopName "
            Case GlobalVariable.ORDERINVENTORY_BYID
                strSQL &= " Order by ShopID "
            Case Else
                strSQL &= " Order by ShopName,ShopCode"
        End Select
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function ListInventoryViewForSelectShop(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal viewFromInventoryID As String, ByVal orderBy As Integer) As DataTable
        Dim strSQL As String
        strSQL = "select ShopID,ShopCode,ShopName " &
                 "from shop_data " &
                 "Where  Deleted = 0 AND IsInv=1 AND ShopID <> " & viewFromInventoryID
        Select Case orderBy
            Case GlobalVariable.ORDERINVENTORY_BYNAME
                strSQL &= " Order by ShopName,ShopCode "
            Case GlobalVariable.ORDERINVENTORY_BYCODEANDNAME
                strSQL &= " Order by ShopCode,ShopName "
            Case GlobalVariable.ORDERINVENTORY_BYID
                strSQL &= " Order by ShopID "
            Case Else
                strSQL &= "Order by ShopName,ShopCode"
        End Select
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetProperty(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection) As DataTable
        Dim strSQL As String = ""
        strSQL = "select * from Property"
        Return dbUtil.List(strSQL, connection)
    End Function
End Module