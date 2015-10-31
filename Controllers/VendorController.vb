Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Public Class VendorController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function CreateVendorGroup(ByVal groupCode As String, ByVal groupName As String, ByRef resultText As String) As Boolean
        Return VendorGroupModule.InsertVendorGroup(globalVariable, groupCode, groupName, resultText)
    End Function

    Public Function UpdateVendorGroup(ByVal groupID As Integer, ByVal groupCode As String, ByVal groupName As String, ByRef resultText As String) As Boolean
        Return VendorGroupModule.UpdateVendorGroup(globalVariable, groupID, groupCode, groupName, resultText)
    End Function

    Public Function DeleteVendorGroup(ByVal groupID As Integer, ByRef resultText As String) As Boolean
        Return VendorGroupModule.DeleteVendorGroup(globalVariable, groupID, resultText)
    End Function

    Public Function ListVendorGroup(ByRef vendorGroupList As List(Of ListVendorGroup_Data), ByRef resultText As String) As Boolean
        Return VendorGroupModule.ListVendorGroup(globalVariable, vendorGroupList, resultText)
    End Function

    Public Function GetVendorGroupDetail(ByVal vendorGroupID As Integer, ByRef vendorGroupData As ListVendorGroup_Data, ByRef resultText As String) As Boolean
        Return VendorGroupModule.GetVendorGroup(globalVariable, vendorGroupID, vendorGroupData, resultText)
    End Function

    Public Function CreateVendor(ByVal groupId As Integer, ByVal vendorCode As String, ByVal vendorName As String, ByVal vendorFirstName As String,
    ByVal vendorLastName As String, ByVal vendorAddress1 As String, ByVal vendorAddress2 As String, ByVal vendorCity As String, ByVal vendorProvince As Integer,
    ByVal vendorZipCode As String, ByVal vendorTelephone As String, ByVal vendorMobile As String, ByVal vendorFax As String, ByVal vendorEmail As String,
    ByVal vendorTermOfPayment As Integer, ByVal vendorCreditDay As Integer, ByVal defaultTaxType As Integer, ByRef resultText As String) As Boolean

        Return VendorModule.InsertVendor(globalVariable, groupId, vendorCode, vendorName, vendorFirstName, vendorLastName,
                                              vendorAddress1, vendorAddress2, vendorCity, vendorProvince, vendorZipCode, vendorTelephone, vendorMobile,
                                              vendorFax, vendorEmail, vendorTermOfPayment, vendorCreditDay, globalVariable.StaffID, defaultTaxType, resultText)
    End Function

    Public Function UpdateVendor(ByVal vendorId As Integer, ByVal groupId As Integer, ByVal vendorCode As String, ByVal vendorName As String,
    ByVal vendorFirstName As String, ByVal vendorLastName As String, ByVal vendorAddress1 As String, ByVal vendorAddress2 As String, ByVal vendorCity As String,
    ByVal vendorProvince As Integer, ByVal vendorZipCode As String, ByVal vendorTelephone As String, ByVal vendorMobile As String, ByVal vendorFax As String,
    ByVal vendorEmail As String, ByVal vendorTermOfPayment As Integer, ByVal vendorCreditDay As Integer, ByVal defaultTaxType As Integer, ByRef resultText As String) As Boolean

        Return VendorModule.UpdateVendor(globalVariable, vendorId, groupId, vendorCode, vendorName, vendorFirstName, vendorLastName,
                                             vendorAddress1, vendorAddress2, vendorCity, vendorProvince, vendorZipCode, vendorTelephone, vendorMobile, vendorFax,
                                             vendorEmail, vendorTermOfPayment, vendorCreditDay, globalVariable.StaffID, defaultTaxType, resultText)
    End Function

    Public Function DeleteVendor(ByVal vendorId As Integer, ByRef resultText As String) As Boolean
        Return VendorModule.DeletedVendor(globalVariable, vendorId, globalVariable.StaffID, resultText)
    End Function

    Public Function ListVendor(ByVal vendorGroupID As Integer, ByRef vendorList As List(Of ListVendorDetail_Data), ByRef resultText As String) As Boolean
        Return VendorModule.ListVendors(globalVariable, vendorGroupID, vendorList, resultText)
    End Function

    Public Function ListVendorDetail(ByVal vendorGroupID As Integer, ByRef vendorList As List(Of VendorFullDetail_Data), ByRef resultText As String) As Boolean
        Return VendorModule.GetListVendorDetail(globalVariable, vendorGroupID, vendorList, resultText)
    End Function

    Public Overloads Function GetVendorDetail(ByVal vendorID As Integer, ByRef vendorData As VendorFullDetail_Data, ByRef resultText As String) As Boolean
        Return VendorModule.GetVendorDetail(globalVariable, vendorID, vendorData, resultText)
    End Function

    Public Function SearchVendorByCodeOrName(ByVal keyWord As String, ByVal searchBy As Integer, ByRef vendorList As List(Of VendorFullDetail_Data),
                                                    ByRef resultText As String) As Boolean
        Return VendorModule.SearchVendorByCodeOrName(globalVariable, keyWord, searchBy, vendorList, resultText)
    End Function

End Class
