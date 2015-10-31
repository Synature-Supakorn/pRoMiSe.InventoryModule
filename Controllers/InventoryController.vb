Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Public Class InventoryController

    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function ListInventoryName(ByRef invList As List(Of ListInventory_Data), ByRef resultText As String) As Boolean
        Return InventoryModule.ListInventoryName(globalVariable, invList, resultText)
    End Function

    Public Function ListInventoryViewForSelectInventory(ByVal viewFromInventoryID As Integer, ByRef invList As List(Of ListInventory_Data), ByRef resultText As String) As Boolean
        Return InventoryModule.ListInventoryViewForSelectInventory(globalVariable, viewFromInventoryID, invList, resultText)
    End Function

End Class
