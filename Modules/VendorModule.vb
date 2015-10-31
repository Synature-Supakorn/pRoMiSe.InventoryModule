Module VendorModule
    Friend Function InsertVendor(ByVal globalVariable As GlobalVariable, ByVal groupId As Integer, ByVal vendorCode As String, ByVal vendorName As String,
                                 ByVal vendorFirstName As String, ByVal vendorLastName As String, ByVal vendorAddress1 As String, ByVal vendorAddress2 As String,
                                 ByVal vendorCity As String, ByVal vendorProvince As Integer, ByVal vendorZipCode As String, ByVal vendorTelephone As String,
                                 ByVal vendorMobile As String, ByVal vendorFax As String, ByVal vendorEmail As String, ByVal vendorTermOfPayment As Integer,
                                 ByVal vendorCreditDay As Integer, ByVal staffId As Integer, ByVal defaultTaxType As Integer, ByRef resultText As String) As Boolean
        Try
            Return VendorSQL.InsertVendor(globalVariable.DocDBUtil, globalVariable.DocConn, groupId, vendorCode, vendorName, vendorFirstName, vendorLastName,
                                             vendorAddress1, vendorAddress2, vendorCity, vendorProvince, vendorZipCode, vendorTelephone, vendorMobile,
                                             vendorFax, vendorEmail, vendorTermOfPayment, vendorCreditDay, staffId, defaultTaxType)
            Return True
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorModule", "InsertVendor", "99", ex.ToString)
            Return False
        End Try
    End Function

    Friend Function UpdateVendor(ByVal globalVariable As GlobalVariable, ByVal vendorId As Integer, ByVal groupId As Integer, ByVal vendorCode As String,
                                        ByVal vendorName As String, ByVal vendorFirstName As String, ByVal vendorLastName As String, ByVal vendorAddress1 As String,
                                        ByVal vendorAddress2 As String, ByVal vendorCity As String, ByVal vendorProvince As Integer, ByVal vendorZipCode As String,
                                        ByVal vendorTelephone As String, ByVal vendorMobile As String, ByVal vendorFax As String, ByVal vendorEmail As String,
                                        ByVal vendorTermOfPayment As Integer, ByVal vendorCreditDay As Integer, ByVal staffId As Integer, ByVal defaultTaxType As Integer,
                                        ByRef resultText As String) As Boolean
        Try
            Return VendorSQL.UpdateVendor(globalVariable.DocDBUtil, globalVariable.DocConn, vendorId, groupId, vendorCode, vendorName, vendorFirstName, vendorLastName,
                                             vendorAddress1, vendorAddress2, vendorCity, vendorProvince, vendorZipCode, vendorTelephone, vendorMobile, vendorFax,
                                             vendorEmail, vendorTermOfPayment, vendorCreditDay, staffId, defaultTaxType)
            Return True
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorModule", "UpdateVendor", "99", ex.ToString)
            Return False
        End Try
    End Function

    Friend Function DeletedVendor(ByVal globalVariable As GlobalVariable, ByVal vendorId As Integer, ByVal staffId As Integer, ByRef resultText As String) As Boolean
        Try
            Return VendorSQL.DeletedVendor(globalVariable.DocDBUtil, globalVariable.DocConn, vendorId, staffId)
            Return True
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorModule", "DeletedVendor", "99", ex.ToString)
            Return False
        End Try
    End Function

    Friend Function ListVendors(ByVal globalVariable As GlobalVariable, ByVal vendorGroupID As Integer, ByRef vendorList As List(Of ListVendorDetail_Data), ByRef resultText As String) As Boolean
        Dim dtVendor As New DataTable
        Dim i As Integer = 0
        Try
            dtVendor = VendorSQL.ListVendors(globalVariable.DocDBUtil, globalVariable.DocConn, vendorGroupID)
            vendorList = New List(Of ListVendorDetail_Data)
            If dtVendor.Rows.Count > 0 Then
                For i = 0 To dtVendor.Rows.Count - 1
                    If IsDBNull(dtVendor.Rows(i)("VendorCode")) Then
                        dtVendor.Rows(i)("VendorCode") = ""
                    End If
                    If IsDBNull(dtVendor.Rows(i)("VendorName")) Then
                        dtVendor.Rows(i)("VendorName") = ""
                    End If
                    If IsDBNull(dtVendor.Rows(i)("MaterialTaxTypeName")) Then
                        dtVendor.Rows(i)("MaterialTaxTypeName") = ""
                    End If
                    vendorList.Add(ListVendorDetail_Data.NewListVendor(dtVendor.Rows(i)("VendorID"), dtVendor.Rows(i)("VendorGroupID"), dtVendor.Rows(i)("VendorCode"), dtVendor.Rows(i)("VendorName"), dtVendor.Rows(i)("MaterialTaxType"), dtVendor.Rows(i)("MaterialTaxTypeName")))
                Next i
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.ToString
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorModule", "ListVendors", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function GetListVendorDetail(ByVal globalVariable As GlobalVariable, ByVal vendorGroupID As Integer, ByRef vendorList As List(Of VendorFullDetail_Data), ByRef resultText As String) As Boolean
        Try
            Dim dtVendor As DataTable
            Dim vendorID As Integer = 0
            vendorList = New List(Of VendorFullDetail_Data)
            dtVendor = VendorSQL.GetListVendorDetail(globalVariable.DocDBUtil, globalVariable.DocConn, vendorGroupID, globalVariable.DocLangID)
            If dtVendor.Rows.Count > 0 Then
                vendorList = VendorFullDetail_Data.NewListVendorFullDetail(dtVendor)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorModule", "GetListVendorDetail", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function GetVendorDetail(ByVal globalVariable As GlobalVariable, ByVal vendorID As Integer, ByRef vendorData As VendorFullDetail_Data, ByRef resultText As String) As Boolean
        Try
            Dim dtVendor As DataTable
            dtVendor = VendorSQL.GetVendorDetail(globalVariable.DocDBUtil, globalVariable.DocConn, vendorID, globalVariable.DocLangID)
            If dtVendor.Rows.Count > 0 Then
                vendorData = VendorFullDetail_Data.NewVendorFullDetail(dtVendor)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorModule", "GetVendorDetail", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function SearchVendorByCodeOrName(ByVal globalVariable As GlobalVariable, ByVal keyWord As String, ByVal searchBy As Integer, ByRef vendorList As List(Of VendorFullDetail_Data),
    ByRef resultText As String) As Boolean
        Try
            Dim dtVendor As New DataTable
            Dim vendorID As Integer = 0
            vendorList = New List(Of VendorFullDetail_Data)
            Select Case searchBy
                Case Is = 1
                    dtVendor = VendorSQL.SearchVendorByCode(globalVariable.DocDBUtil, globalVariable.DocConn, keyWord, globalVariable.DocLangID)
                Case Is = 2
                    dtVendor = VendorSQL.SearchVendorByName(globalVariable.DocDBUtil, globalVariable.DocConn, keyWord, globalVariable.DocLangID)
            End Select
            If dtVendor.Rows.Count > 0 Then
                vendorList = VendorFullDetail_Data.NewListVendorFullDetail(dtVendor)
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorModule", "SearchVendorByCodeOrName", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

End Module
