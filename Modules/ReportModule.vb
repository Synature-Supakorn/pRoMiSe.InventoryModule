Imports pRoMiSe.Utilitys.Utilitys
Imports System.Data.SqlClient
Module ReportModule

    Friend Function SearchDocument(ByVal globalVariable As GlobalVariable, ByVal documentTypeId As Integer, ByVal documentStatus As Integer, ByVal startDate As Date, ByVal endDate As Date,
                                  ByVal searchInventoryID As Integer, ByVal vendorID As Integer, ByVal vendorGroupID As Integer,
                                  ByRef docList As List(Of SearchDocumentResult_Data), ByRef resultText As String) As Boolean

        Dim dtResult, dtDocStatus As DataTable
        Dim strFromDate, strToDate As String
        If startDate = Date.MinValue Then
            strFromDate = ""
        Else
            strFromDate = FormatDate(startDate)
        End If
        If endDate = Date.MinValue Then
            strToDate = ""
        Else
            strToDate = FormatDate(endDate)
        End If

        Try
            dtResult = DocumentSQL.SearchDocument(globalVariable.DocDBUtil, globalVariable.DocConn, documentTypeId, strFromDate, strToDate,
                                                       documentStatus, searchInventoryID, vendorID, vendorGroupID, globalVariable.DocLangID)
            dtDocStatus = DocumentSQL.SearchStatusDocument(globalVariable.DocDBUtil, globalVariable.DocConn)
            If dtResult.Rows.Count > 0 Then
                docList = DocumentModule.InsertResultDataIntoList(globalVariable, documentTypeId, dtResult, dtDocStatus)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            Return False
        End Try
        resultText = ""
        Return True
    End Function

End Module
