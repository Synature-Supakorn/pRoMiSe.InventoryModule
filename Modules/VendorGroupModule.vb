Module VendorGroupModule

    Friend Function InsertVendorGroup(ByVal globalVariable As GlobalVariable, ByVal groupCode As String, ByVal groupName As String, ByRef resultText As String) As Boolean
        Try
            Return VendorSQL.InsertVendorGroup(globalVariable.DocDBUtil, globalVariable.DocConn, groupCode, groupName)
            Return True
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorGroupModule", "InsertVendorGroup", "99", ex.ToString)
            Return False
        End Try
    End Function

    Friend Function UpdateVendorGroup(ByVal globalVariable As GlobalVariable, ByVal groupID As Integer, ByVal groupCode As String, ByVal groupName As String, ByRef resultText As String) As Boolean
        Try
            Return VendorSQL.UpdateVendorGroup(globalVariable.DocDBUtil, globalVariable.DocConn, groupID, groupCode, groupName)
            Return True
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorGroupModule", "UpdateVendorGroup", "99", ex.ToString)
            Return False
        End Try
    End Function

    Friend Function DeleteVendorGroup(ByVal globalVariable As GlobalVariable, ByVal groupID As Integer, ByRef resultText As String) As Boolean
        Try
            Return VendorSQL.DeleteVendorGroup(globalVariable.DocDBUtil, globalVariable.DocConn, groupID)
            Return True
        Catch ex As Exception
            resultText = ex.Message
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorGroupModule", "DeleteVendorGroup", "99", ex.ToString)
            Return False
        End Try
    End Function

    Friend Function ListVendorGroup(ByVal globalVariable As GlobalVariable, ByRef vendorGroupList As List(Of ListVendorGroup_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtVendorGroup As DataTable
        Dim strCanViewGroupID As String = ""
        Try
            dtVendorGroup = VendorSQL.ListVendorGroup(globalVariable.DocDBUtil, globalVariable.DocConn)
            vendorGroupList = New List(Of ListVendorGroup_Data)

            If dtVendorGroup.Rows.Count > 0 Then
                For i = 0 To dtVendorGroup.Rows.Count - 1
                    If IsDBNull(dtVendorGroup.Rows(i)("VendorGroupCode")) Then
                        dtVendorGroup.Rows(i)("VendorGroupCode") = ""
                    End If
                    If IsDBNull(dtVendorGroup.Rows(i)("VendorGroupName")) Then
                        dtVendorGroup.Rows(i)("VendorGroupName") = ""
                    End If
                    vendorGroupList.Add(ListVendorGroup_Data.NewListVendorGroup(dtVendorGroup.Rows(i)("VendorGroupID"),
                    dtVendorGroup.Rows(i)("VendorGroupCode"), dtVendorGroup.Rows(i)("VendorGroupName")))
                Next i
            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If

        Catch ex As Exception
            resultText = ex.ToString
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorGroupModule", "ListVendorGroup", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function GetVendorGroup(ByVal globalVariable As GlobalVariable, ByVal vendorGroupID As Integer, ByRef vendorGroupData As ListVendorGroup_Data, ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtVendorGroup As DataTable
        Dim strCanViewGroupID As String = ""
        Try
            dtVendorGroup = VendorSQL.GetVendorGroup(globalVariable.DocDBUtil, globalVariable.DocConn, vendorGroupID)
            vendorGroupData = New ListVendorGroup_Data
            If dtVendorGroup.Rows.Count > 0 Then

                If IsDBNull(dtVendorGroup.Rows(0)("VendorGroupCode")) Then
                    dtVendorGroup.Rows(0)("VendorGroupCode") = ""
                End If
                If IsDBNull(dtVendorGroup.Rows(0)("VendorGroupName")) Then
                    dtVendorGroup.Rows(0)("VendorGroupName") = ""
                End If
                vendorGroupData = ListVendorGroup_Data.NewListVendorGroup(dtVendorGroup.Rows(0)("VendorGroupID"),
                                dtVendorGroup.Rows(0)("VendorGroupCode"), dtVendorGroup.Rows(0)("VendorGroupName"))

            Else
                resultText = globalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.ToString
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "VendorGroupModule", "GetVendorGroup", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

End Module
