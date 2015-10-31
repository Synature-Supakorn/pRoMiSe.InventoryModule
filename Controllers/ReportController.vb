Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Public Class ReportController
    Private globalVariable As New GlobalVariable

    Sub New(ByVal dbUtil As CDBUtil, ByVal conn As SqlConnection, ByVal staffID As Integer, ByVal langID As Integer)
        globalVariable.DocDBUtil = dbUtil
        globalVariable.DocConn = conn
        globalVariable.DocLangID = langID
        globalVariable.StaffID = staffID
        InventoryModule.GetProperty(globalVariable)
    End Sub

    Public Function TaxReportOil(ByVal selMonth As Integer, ByVal selYear As Integer, ByVal inventoryID As Integer,
                                 ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean
        Dim startDate As Date
        Dim endDate As Date
        startDate = New Date(selYear, selMonth, 1)
        endDate = New Date(selYear, selMonth, DateTime.DaysInMonth(selYear, selMonth))
        Return ReportModule.SearchDocument(globalVariable, globalVariable.DOCUMENTTYPE_DIRECTROPTT, globalVariable.DOCUMENTSTATUS_APPROVE, startDate, endDate, inventoryID, -1, -1, docList, resultText)
    End Function

    Public Function TaxReportNonOil(ByVal selMonth As Integer, ByVal selYear As Integer, ByVal inventoryID As Integer,
                                    ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean
        Dim startDate As Date
        Dim endDate As Date
        startDate = New Date(selYear, selMonth, 1)
        endDate = New Date(selYear, selMonth, DateTime.DaysInMonth(selYear, selMonth))
        Return ReportModule.SearchDocument(globalVariable, globalVariable.DOCUMENTTYPE_DIRECTROPTTNONOIL, globalVariable.DOCUMENTSTATUS_APPROVE, startDate, endDate, inventoryID, -1, -1, docList, resultText)
    End Function

End Class
