Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient
Imports pRoMiSe.Utilitys.Utilitys

Public Class StockCardController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Function StockCardReport(ByVal inventoryId As Integer, ByVal materialGroupId As Integer, ByVal materialDeptId As Integer,
                             ByVal materialCode As String, ByVal startDate As Date, ByVal endDate As Date,
                             ByRef stockCardData As List(Of StockCard_Data), ByRef resultText As String) As Boolean

        Return StockCardModule.GetStockCardReport(globalVariable, inventoryId, materialGroupId, materialDeptId, materialCode,
                                                  startDate, endDate, stockCardData, resultText)
    End Function

End Class
