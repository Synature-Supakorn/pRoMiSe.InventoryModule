Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Public Class MaterialController

    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function ListMaterialGroup(ByRef materialGroupData As List(Of ListMaterialGroup_Data), ByRef resultText As String) As Boolean
        Return MaterialGroupModule.ListMaterialGroup(globalVariable, materialGroupData, resultText)
    End Function

    Public Function ListMaterialDept(ByVal materialGroupID As Integer, ByRef materialDeptData As List(Of ListMaterialDept_Data), ByRef resultText As String) As Boolean
        Return MaterialDeptModule.ListMaterialDept(globalVariable, materialGroupID, materialDeptData, resultText)
    End Function

    Public Function ListMaterial(ByVal materialGroupID As Integer, ByVal materialDeptID As Integer, ByRef materialList As List(Of ListMaterialDetail_Data), ByRef resultText As String) As Boolean

        Dim dtMaterialDefaultPrice As New DataTable
        Dim dtMaterials As New DataTable
        Dim dtUnits As New DataTable
        Dim defaultTaxType As Integer = 99

        dtMaterials = MaterialSQL.ListMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, materialGroupID, materialDeptID)
        dtUnits = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn)
        Return MaterialModule.InsertMaterialFromDataTableToList(globalVariable, dtMaterials, dtUnits, dtMaterialDefaultPrice, defaultTaxType, materialList, resultText)
    End Function

    Public Function ListMaterial(ByVal materialGroupID As Integer, ByVal materialDeptID As Integer, ByVal vendorId As Integer, ByRef materialList As List(Of ListMaterialDetail_Data), ByRef resultText As String) As Boolean

        Dim dtMaterialDefaultPrice As New DataTable
        Dim dtMaterials As New DataTable
        Dim dtUnits As New DataTable
        Dim dtVendor As New DataTable
        Dim defaultTaxType As Integer = 1

        dtVendor = VendorSQL.GetVendorDetail(globalVariable.DocDBUtil, globalVariable.DocConn, vendorId, globalVariable.DocLangID)
        If dtVendor.Rows.Count > 0 Then
            defaultTaxType = dtVendor.Rows(0)("defaulttaxtype")
        End If
        dtMaterialDefaultPrice = MaterialSQL.GetMaterialDefaultPrice(globalVariable.DocDBUtil, globalVariable.DocConn, vendorId)
        dtMaterials = MaterialSQL.ListMaterial(globalVariable.DocDBUtil, globalVariable.DocConn, materialGroupID, materialDeptID)
        dtUnits = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn)
        Return MaterialModule.InsertMaterialFromDataTableToList(globalVariable, dtMaterials, dtUnits, dtMaterialDefaultPrice, defaultTaxType, materialList, resultText)
    End Function

    Public Function GetMaterialByCode(ByVal materialCode As String, ByRef materialList As List(Of ListMaterialDetail_Data), ByRef resultText As String) As Boolean
        Dim dtMaterials As New DataTable
        Dim dtUnits As New DataTable
        Dim dtMaterialDefaultPrice As New DataTable
        Dim defaultTaxType As Integer = 99
        dtMaterials = MaterialSQL.GetMaterialDetailFromCode(globalVariable.DocDBUtil, globalVariable.DocConn, materialCode, False)
        dtUnits = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn)
        Return MaterialModule.InsertMaterialFromDataTableToList(globalVariable, dtMaterials, dtUnits, dtMaterialDefaultPrice, defaultTaxType, materialList, resultText)
    End Function

    Public Function GetMaterialByCode(ByVal materialCode As String, ByVal vendorID As Integer, ByRef materialList As List(Of ListMaterialDetail_Data), ByRef resultText As String) As Boolean
        Dim dtMaterials As New DataTable
        Dim dtUnits As New DataTable
        Dim dtMaterialDefaultPrice As New DataTable
        Dim dtVendor As New DataTable
        Dim defaultTaxType As Integer = 1
        dtVendor = VendorSQL.GetVendorDetail(globalVariable.DocDBUtil, globalVariable.DocConn, vendorID, globalVariable.DocLangID)
        If dtVendor.Rows.Count > 0 Then
            defaultTaxType = dtVendor.Rows(0)("defaulttaxtype")
        End If
        dtMaterialDefaultPrice = MaterialSQL.GetMaterialDefaultPrice(globalVariable.DocDBUtil, globalVariable.DocConn, vendorID)
        dtMaterials = MaterialSQL.GetMaterialDetailFromCode(globalVariable.DocDBUtil, globalVariable.DocConn, materialCode, False)
        dtUnits = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn)
        Return MaterialModule.InsertMaterialFromDataTableToList(globalVariable, dtMaterials, dtUnits, dtMaterialDefaultPrice, defaultTaxType, materialList, resultText)
    End Function

    Public Function GetMaterialByID(ByVal materialID As Integer, ByRef materialList As List(Of ListMaterialDetail_Data), ByRef resultText As String) As Boolean
        Dim dtMaterials As New DataTable
        Dim dtUnits As New DataTable
        Dim dtMaterialDefaultPrice As New DataTable
        Dim defaultTaxType As Integer = 99
        dtMaterials = MaterialSQL.GetMaterialDetail(globalVariable.DocDBUtil, globalVariable.DocConn, materialID)
        dtUnits = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn)
        Return MaterialModule.InsertMaterialFromDataTableToList(globalVariable, dtMaterials, dtUnits, dtMaterialDefaultPrice, defaultTaxType, materialList, resultText)
    End Function

    Public Function GetMaterialByID(ByVal materialID As Integer, ByVal vendorId As Integer, ByRef materialList As List(Of ListMaterialDetail_Data), ByRef resultText As String) As Boolean
        Dim dtMaterials As New DataTable
        Dim dtUnits As New DataTable
        Dim dtMaterialDefaultPrice As New DataTable
        Dim dtVendor As New DataTable
        Dim defaultTaxType As Integer = 1
        dtVendor = VendorSQL.GetVendorDetail(globalVariable.DocDBUtil, globalVariable.DocConn, vendorId, globalVariable.DocLangID)
        If dtVendor.Rows.Count > 0 Then
            defaultTaxType = dtVendor.Rows(0)("defaulttaxtype")
        End If
        dtMaterialDefaultPrice = MaterialSQL.GetMaterialDefaultPrice(globalVariable.DocDBUtil, globalVariable.DocConn, vendorId)
        dtMaterials = MaterialSQL.GetMaterialDetail(globalVariable.DocDBUtil, globalVariable.DocConn, materialID)
        dtUnits = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn)
        Return MaterialModule.InsertMaterialFromDataTableToList(globalVariable, dtMaterials, dtUnits, dtMaterialDefaultPrice, defaultTaxType, materialList, resultText)
    End Function

    Public Function SearchMaterialByCodeOrName(ByVal keyWord As String, ByVal isSearchCode As Boolean, ByRef materialList As List(Of ListMaterialDetail_Data), ByRef resultText As String) As Boolean
        Dim materialGroupType As String = ""
        Dim vendorId As Integer = 0
        Dim listMaterialBy As Integer = 0
        Return MaterialModule.SearchMaterialByCodeOrName(globalVariable, keyWord, isSearchCode, materialList, resultText)
    End Function

    Public Function SearchMaterialByCodeOrName(ByVal keyWord As String, ByVal isSearchCode As Boolean, ByVal vendorId As Integer, ByRef materialList As List(Of ListMaterialDetail_Data), ByRef resultText As String) As Boolean
        Dim materialGroupType As String = ""
        Dim listMaterialBy As Integer = 0
        Return MaterialModule.SearchMaterialByCodeOrName(globalVariable, keyWord, isSearchCode, materialList, resultText)
    End Function

    Public Function GetMaterialUnit(ByVal materialID As Integer, ByRef materialUnitList As List(Of ListMaterialUnit_Data), ByRef resultText As String) As Boolean
        Dim dtUnits As New DataTable
        Dim dtMaterialDefaultPrice As New DataTable
        dtUnits = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn, materialID)
        Return MaterialModule.InsertMaterialUnitFromDataTableToList(globalVariable, dtUnits, dtMaterialDefaultPrice, materialUnitList, resultText)
    End Function

    Public Function ListMaterialUnitLarge(ByRef materialUnitList As List(Of ListMaterialUnit_Data), ByRef resultText As String) As Boolean
        Dim dtUnits As New DataTable
        Dim dtMaterialDefaultPrice As New DataTable
        dtUnits = MaterialSQL.ListMaterialUnitLarge(globalVariable.DocDBUtil, globalVariable.DocConn)
        Return MaterialModule.InsertMaterialUnitFromDataTableToList(globalVariable, dtUnits, dtMaterialDefaultPrice, materialUnitList, resultText)
    End Function

    Public Function ListMaterialTaxType(ByRef taxTypeList As List(Of ListMaterialTaxType_Data), ByRef resultText As String) As Boolean
        Try
            Dim dtTaxType As New DataTable
            Dim strInclude As String = "0,1,2"
            dtTaxType = MaterialSQL.GetMaterialTaxType(globalVariable.DocDBUtil, globalVariable.DocConn, strInclude)
            taxTypeList = ListMaterialTaxType_Data.ListMaterialTaxType(dtTaxType)
        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function ListMaterialTaxType(ByVal vendorId As Integer, ByRef taxTypeList As List(Of ListMaterialTaxType_Data), ByRef resultText As String) As Boolean
        Try
            Dim dt As New DataTable
            Dim dtTaxType As New DataTable
            Dim strInclude As String = "0,1"
            Dim strExclude As String = "0,2"
            dt = VendorSQL.GetVendorDetail(globalVariable.DocDBUtil, globalVariable.DocConn, vendorId, globalVariable.DocLangID)
            If dt.Rows.Count > 0 Then
                Select Case dt.Rows(0)("defaulttaxtype")
                    Case Is = globalVariable.TAXTYPE_INCLUDEVAT
                        dtTaxType = MaterialSQL.GetMaterialTaxType(globalVariable.DocDBUtil, globalVariable.DocConn, strInclude)
                    Case Is = globalVariable.TAXTYPE_EXCLUDEVAT
                        dtTaxType = MaterialSQL.GetMaterialTaxType(globalVariable.DocDBUtil, globalVariable.DocConn, strExclude)
                    Case Else
                        dtTaxType = MaterialSQL.GetMaterialTaxType(globalVariable.DocDBUtil, globalVariable.DocConn, strInclude)
                End Select
                taxTypeList = ListMaterialTaxType_Data.ListMaterialTaxType(dtTaxType)
            Else
                dtTaxType = MaterialSQL.GetMaterialTaxType(globalVariable.DocDBUtil, globalVariable.DocConn, strInclude)
                taxTypeList = ListMaterialTaxType_Data.ListMaterialTaxType(dtTaxType)
            End If

        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Public Function ListMaterialDiscountType(ByRef discountType As List(Of ListMaterialDiscountType_Data), ByRef resultText As String) As Boolean
        Try
            discountType = ListMaterialDiscountType_Data.ListMaterialDiscountType
        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function
End Class
