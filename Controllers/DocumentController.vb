Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Public Class DocumentController

    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function StatusDocument(ByRef statusDocumentList As List(Of Status_Data), ByRef resultText As String) As Boolean
        Return DocumentModule.StatusDocument(globalVariable, statusDocumentList, resultText)
    End Function

    Public Function LoadDocumentDetail(ByVal documentID As Integer, ByVal documentShopID As Integer, ByVal oldDocumentID As Integer,
                                                ByVal oldDocumentShopID As Integer, ByRef documentData As Document_Data, ByRef resultText As String) As Boolean
        Return DocumentModule.LoadDocument(globalVariable, documentID, documentShopID, documentData, resultText)
    End Function
   
End Class
