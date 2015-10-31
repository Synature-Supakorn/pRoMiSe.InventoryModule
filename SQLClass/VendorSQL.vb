Imports pRoMiSe.DBHelper
Imports pRoMiSe.Utilitys.Utilitys
Imports System.Data.SqlClient

Module VendorSQL

    Friend Function InsertVendorGroup(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal groupCode As String, ByVal groupName As String) As Integer
        Dim strSQL As String
        Dim dt As New DataTable
        Dim groupId As Integer = 0

        dt = dbUtil.List("Select isnull(max(VendorGroupID),1)As VendorGroupId From vendorgroup", connection)
        groupId = dt.Rows(0)("VendorGroupId") + 1

        strSQL = "Insert Into vendorgroup(VendorGroupID,VendorGroupCode,VendorGroupName,Deleted)" &
                 "Values(" & groupId & ",'" & Replace(groupCode) & "','" & Replace(groupName) & "',0)"
        Return dbUtil.sqlExecute(strSQL, connection)
    End Function

    Friend Function UpdateVendorGroup(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal groupId As Integer, ByVal groupCode As String, ByVal groupName As String) As Integer
        Dim strSQL As String
        Dim dt As New DataTable
        strSQL = "Update vendorgroup Set VendorGroupID=" & groupId &
                 ",VendorGroupCode='" & Replace(groupCode) &
                 "',VendorGroupName='" & Replace(groupName) &
                 "' Where VendorGroupID =" & groupId
        Return dbUtil.sqlExecute(strSQL, connection)
    End Function

    Friend Function DeleteVendorGroup(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal groupId As Integer) As Integer
        Dim strSQL As String
        Dim dt As New DataTable
        strSQL = "Update vendorgroup Set Deleted=1 " &
                 " Where VendorGroupID =" & groupId
        Return dbUtil.sqlExecute(strSQL, connection)
    End Function

    Friend Function ListVendorGroup(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection) As DataTable
        Dim strSQL As String
        strSQL = " Select VendorGroupID, VendorGroupCode, VendorGroupName " &
                 " From VendorGroup " &
                 " Where Deleted = 0" &
                 " Order by VendorGroupName "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetVendorGroup(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal vendorGroupID As Integer) As DataTable
        Dim strSQL As String
        strSQL = " Select VendorGroupID, VendorGroupCode, VendorGroupName " &
                 " From VendorGroup " &
                 " Where Deleted = 0 AND VendorGroupID=" & vendorGroupID &
                 " Order by VendorGroupName "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function InsertVendor(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal groupId As Integer, ByVal vendorCode As String,
                                 ByVal vendorName As String, ByVal vendorFirstName As String, ByVal vendorLastName As String, ByVal vendorAddress1 As String,
                                 ByVal vendorAddress2 As String, ByVal vendorCity As String, ByVal vendorProvince As Integer, ByVal vendorZipCode As String,
                                 ByVal vendorTelephone As String, ByVal vendorMobile As String, ByVal vendorFax As String, ByVal vendorEmail As String,
                                 ByVal vendorTermOfPayment As Integer, ByVal vendorCreditDay As Integer, ByVal staffId As Integer, ByVal defaultTaxType As Integer) As Integer

        Dim strSQL As String
        Dim dt As New DataTable
        Dim currentDate As DateTime = Now
        Dim vendorId As Integer = 0

        dt = dbUtil.List("Select isnull(max(VendorID),1)As VendorId From vendors", connection)
        vendorId = dt.Rows(0)("VendorId") + 1

        strSQL = "Insert Into Vendors(VendorID,VendorGroupID,VendorName,VendorCode,VendorFirstName,VendorLastName,VendorAddress1,VendorAddress2," &
                 "VendorCity,VendorProvince,VendorZipCode,VendorTelephone,VendorMobile,VendorFax,VendorEmail,VendorTermOfPayment,VendorCreditDay,InsertDate,InputBy,DefaultTaxType)" &
                 "values(" & vendorId & "," & groupId & ",'" & Replace(vendorName) & "','" & Replace(vendorCode) & "'" &
                 ",'" & Replace(vendorFirstName) & "','" & Replace(vendorLastName) & "','" & Replace(vendorAddress1) & "'" &
                 ",'" & Replace(vendorAddress2) & "','" & Replace(vendorCity) & "'," & vendorProvince &
                 ",'" & Replace(vendorZipCode) & "','" & Replace(vendorTelephone) & "','" & Replace(vendorMobile) & "'" &
                 ",'" & Replace(vendorFax) & "','" & Replace(vendorEmail) & "','" & Replace(vendorTermOfPayment) & "'" &
                 "," & vendorCreditDay & "," & FormatDateTime(currentDate) & "," & staffId & "," & defaultTaxType & ")"
        Return dbUtil.sqlExecute(strSQL, connection)
    End Function

    Friend Function UpdateVendor(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal vendorId As Integer, ByVal groupId As Integer,
                                 ByVal vendorCode As String, ByVal vendorName As String, ByVal vendorFirstName As String, ByVal vendorLastName As String,
                                 ByVal vendorAddress1 As String, ByVal vendorAddress2 As String, ByVal vendorCity As String, ByVal vendorProvince As Integer,
                                 ByVal vendorZipCode As String, ByVal vendorTelephone As String, ByVal vendorMobile As String, ByVal vendorFax As String,
                                 ByVal vendorEmail As String, ByVal vendorTermOfPayment As Integer, ByVal vendorCreditDay As Integer, ByVal staffId As Integer,
                                 ByVal defaultTaxType As Integer) As Integer

        Dim strSQL As String
        Dim dt As New DataTable
        Dim currentDate As DateTime = Now

        strSQL = " Update Vendors" &
                 " Set VendorGroupID=" & groupId & ",VendorName='" & Replace(vendorName) & "',VendorCode='" & Replace(vendorCode) & "'" &
                 ",VendorFirstName='" & Replace(vendorFirstName) & "',VendorLastName='" & Replace(vendorLastName) & "',VendorAddress1='" & Replace(vendorAddress1) & "'" &
                 ",VendorAddress2='" & Replace(vendorAddress2) & "',VendorCity='" & Replace(vendorCity) & "',VendorProvince=" & vendorProvince &
                 ",VendorZipCode='" & Replace(vendorZipCode) & "',VendorTelephone='" & Replace(vendorTelephone) & "',VendorMobile='" & Replace(vendorMobile) & "'" &
                 ",VendorFax='" & Replace(vendorFax) & "',VendorEmail='" & Replace(vendorEmail) & "',VendorTermOfPayment='" & Replace(vendorTermOfPayment) & "'" &
                 ",VendorCreditDay=" & vendorCreditDay & ",UpdateDate=" & FormatDateTime(currentDate) & ",InputBy=" & staffId & ",DefaultTaxType=" & defaultTaxType & " Where VendorID=" & vendorId
        Return dbUtil.sqlExecute(strSQL, connection)
    End Function

    Friend Function DeletedVendor(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal vendorId As Integer, ByVal staffId As Integer) As Integer
        Dim strSQL As String
        Dim dt As New DataTable
        Dim currentDate As DateTime = Now
        strSQL = " Update Vendors" &
                 " Set Deleted=1,UpdateDate=" & FormatDateTime(currentDate) & " Where VendorID=" & vendorId
        Return dbUtil.sqlExecute(strSQL, connection)
    End Function

    Friend Function ListVendors(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal groupID As Integer) As DataTable
        Dim strSQL As String
        strSQL = " Select v.VendorID, v.VendorGroupID, v.VendorCode, v.VendorName, t.*" &
                 " From Vendors v " &
                 " LEFT OUTER JOIN MaterialTaxtype t ON v.defaultTaxType=t.MaterialTaxType " &
                 " Where v.VendorGroupID = " & groupID & " And d.Deleted = 0 " &
                 " Order by v.VendorCode, v.VendorName "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetListVendorDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal vendorGroupID As Integer, ByVal langID As Integer) As DataTable
        Dim strSQL As String
        strSQL = " Select vg.VendorGroupCode, vg.VendorGroupName, v.*, p.ProvinceName, t.* " &
                 " From VendorGroup vg, Vendors v LEFT OUTER JOIN Provinces p " &
                 " ON p.ProvinceID = v.VendorProvince AND p.LangID = " & langID &
                 " LEFT OUTER JOIN MaterialTaxtype t ON v.defaultTaxType=t.MaterialTaxType " &
                 " Where v.Deleted=0 AND v.VendorGroupID = vg.VendorGroupID  AND v.VendorGroupID = " & vendorGroupID
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function GetVendorDetail(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal vendorID As Integer, ByVal langID As Integer) As DataTable
        Dim strSQL As String
        strSQL = " Select vg.VendorGroupCode, vg.VendorGroupName, v.*, p.ProvinceName, t.* " &
                 " From VendorGroup vg, Vendors v LEFT OUTER JOIN Provinces p " &
                 " ON p.ProvinceID = v.VendorProvince AND p.LangID = " & langID &
                 " LEFT OUTER JOIN MaterialTaxtype t ON v.defaultTaxType=t.MaterialTaxType " &
                 " Where v.VendorGroupID = vg.VendorGroupID AND v.VendorID=" & vendorID
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function SearchVendorByCode(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal keyWord As String, ByVal langID As Integer) As DataTable
        Dim strSQL As String
        strSQL = " Select vg.VendorGroupCode, vg.VendorGroupName, v.*, p.ProvinceName, t.* " &
                 " From VendorGroup vg, Vendors v LEFT OUTER JOIN Provinces p " &
                 " ON p.ProvinceID = v.VendorProvince AND p.LangID = " & langID &
                 " LEFT OUTER JOIN MaterialTaxtype t ON v.defaultTaxType=t.MaterialTaxType " &
                 " Where  vg.VendorGroupID = v.VendorGroupID AND  v.VendorCode Like '%" & keyWord & "%' AND v.Deleted = 0 " &
                 " Order by v.VendorCode, v.VendorName "
        Return dbUtil.List(strSQL, connection)
    End Function

    Friend Function SearchVendorByName(ByVal dbUtil As CDBUtil, ByVal connection As SqlConnection, ByVal keyWord As String, ByVal langID As Integer) As DataTable
        Dim strSQL As String
        strSQL = " Select vg.VendorGroupCode, vg.VendorGroupName, v.*, p.ProvinceName, t.* " &
                " From VendorGroup vg, Vendors v LEFT OUTER JOIN Provinces p " &
                " ON p.ProvinceID = v.VendorProvince AND p.LangID = " & langID &
                " LEFT OUTER JOIN MaterialTaxtype t ON v.defaultTaxType=t.MaterialTaxType " &
                " Where  vg.VendorGroupID = v.VendorGroupID AND  v.VendorName Like '%" & keyWord & "%' AND v.Deleted = 0 " &
                " Order by v.VendorCode, v.VendorName "
        Return dbUtil.List(strSQL, connection)
    End Function

End Module
