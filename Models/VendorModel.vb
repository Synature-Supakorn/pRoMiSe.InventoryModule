Public Class ListVendorDetail_Data
    Public VendorID As Integer
    Public VendorGroupID As Integer
    Public VendorShopID As Integer
    Public VendorCode As String
    Public VendorName As String
    Public DefaultTaxType As Integer
    Public DefaultTaxTypeName As String

    Public Shared Function NewListVendor(ByVal vendorID As Integer, ByVal vendorGroupID As Integer, ByVal vendorCode As String, ByVal vendorName As String,
                                         ByVal defaultTaxType As Integer, ByVal defaultTaxTypeName As String) As ListVendorDetail_Data

        Dim vData As New ListVendorDetail_Data
        vData.VendorID = vendorID
        vData.VendorGroupID = vendorGroupID
        vData.VendorShopID = 0
        vData.VendorCode = vendorCode
        vData.VendorName = vendorName
        vData.DefaultTaxType = defaultTaxType
        vData.DefaultTaxTypeName = defaultTaxTypeName
        Return vData
    End Function
End Class

Public Class VendorFullDetail_Data
    Public VendorID As Integer
    Public VendorGroupID As Integer
    Public VendorShopID As Integer
    Public VendorGroupCode As String
    Public VendorGroupName As String
    Public VendorCode As String
    Public VendorName As String
    Public VendorFirstName As String
    Public VendorLastName As String
    Public VendorAddress1 As String
    Public VendorAddress2 As String
    Public VendorCity As String
    Public VendorProvinceID As Integer
    Public VendorProvinceName As String
    Public VendorZipCode As String
    Public VendorTelephone As String
    Public VendorFax As String
    Public VendorMobile As String
    Public VendorEMail As String
    Public VendorAddtional As String
    Public TermOfPayment As Integer
    Public CreditDay As Integer
    Public DefaultTaxType As Integer
    Public DefaultTaxTypeName As String

    Public Shared Function NewVendorFullDetail(ByVal dtVendor As DataTable) As VendorFullDetail_Data
        Dim vData As New VendorFullDetail_Data
        If dtVendor.Rows.Count > 0 Then
            vData.VendorID = dtVendor.Rows(0)("VendorID")
            vData.VendorGroupID = dtVendor.Rows(0)("VendorGroupID")
            vData.VendorCode = dtVendor.Rows(0)("VendorCode")
            vData.VendorName = dtVendor.Rows(0)("VendorName")
            vData.VendorGroupCode = dtVendor.Rows(0)("VendorGroupCode")
            vData.VendorGroupName = dtVendor.Rows(0)("VendorGroupName")
            If IsDBNull(dtVendor.Rows(0)("VendorFirstName")) Then
                dtVendor.Rows(0)("VendorFirstName") = ""
            End If
            vData.VendorFirstName = dtVendor.Rows(0)("VendorFirstName")
            If IsDBNull(dtVendor.Rows(0)("VendorLastName")) Then
                dtVendor.Rows(0)("VendorLastName") = ""
            End If
            vData.VendorLastName = dtVendor.Rows(0)("VendorLastName")
            If IsDBNull(dtVendor.Rows(0)("VendorAddress1")) Then
                dtVendor.Rows(0)("VendorAddress1") = ""
            End If
            vData.VendorAddress1 = dtVendor.Rows(0)("VendorAddress1")
            If IsDBNull(dtVendor.Rows(0)("VendorAddress2")) Then
                dtVendor.Rows(0)("VendorAddress2") = ""
            End If
            vData.VendorAddress2 = dtVendor.Rows(0)("VendorAddress2")
            If IsDBNull(dtVendor.Rows(0)("VendorCity")) Then
                dtVendor.Rows(0)("VendorCity") = ""
            End If
            vData.VendorCity = dtVendor.Rows(0)("VendorCity")
            If IsDBNull(dtVendor.Rows(0)("ProvinceName")) Then
                dtVendor.Rows(0)("ProvinceName") = ""
            End If
            If IsDBNull(dtVendor.Rows(0)("VendorProvince")) Then
                dtVendor.Rows(0)("VendorProvince") = 0
            End If
            vData.VendorProvinceID = dtVendor.Rows(0)("VendorProvince")
            vData.VendorProvinceName = dtVendor.Rows(0)("ProvinceName")
            If IsDBNull(dtVendor.Rows(0)("VendorZipCode")) Then
                dtVendor.Rows(0)("VendorZipCode") = ""
            End If
            vData.VendorZipCode = dtVendor.Rows(0)("VendorZipCode")
            If IsDBNull(dtVendor.Rows(0)("VendorTelephone")) Then
                dtVendor.Rows(0)("VendorTelephone") = ""
            End If
            vData.VendorTelephone = dtVendor.Rows(0)("VendorTelephone")
            If IsDBNull(dtVendor.Rows(0)("VendorFax")) Then
                dtVendor.Rows(0)("VendorFax") = ""
            End If
            vData.VendorFax = dtVendor.Rows(0)("VendorFax")
            If IsDBNull(dtVendor.Rows(0)("VendorMobile")) Then
                dtVendor.Rows(0)("VendorMobile") = ""
            End If
            vData.VendorMobile = dtVendor.Rows(0)("VendorMobile")
            If IsDBNull(dtVendor.Rows(0)("VendorEmail")) Then
                dtVendor.Rows(0)("VendorEmail") = ""
            End If
            vData.VendorEMail = dtVendor.Rows(0)("VendorEmail")
            If IsDBNull(dtVendor.Rows(0)("VendorAdditional")) Then
                dtVendor.Rows(0)("VendorAdditional") = ""
            End If
            vData.VendorAddtional = dtVendor.Rows(0)("VendorAdditional")
            If IsDBNull(dtVendor.Rows(0)("VendorTermOfPayment")) Then
                dtVendor.Rows(0)("VendorTermOfPayment") = 0
            End If
            vData.TermOfPayment = dtVendor.Rows(0)("VendorTermOfPayment")
            If IsDBNull(dtVendor.Rows(0)("VendorCreditDay")) Then
                dtVendor.Rows(0)("VendorCreditDay") = 0
            End If
            vData.CreditDay = dtVendor.Rows(0)("VendorCreditDay")
            vData.DefaultTaxType = dtVendor.Rows(0)("MaterialTaxType")
            If Not IsDBNull(dtVendor.Rows(0)("MaterialTaxTypeName")) Then
                vData.DefaultTaxTypeName = dtVendor.Rows(0)("MaterialTaxTypeName")
            End If

        End If

        Return vData
    End Function
    Public Shared Function NewListVendorFullDetail(ByVal dtVendor As DataTable) As List(Of VendorFullDetail_Data)

        Dim vData As VendorFullDetail_Data
        Dim vListData As New List(Of VendorFullDetail_Data)

        If dtVendor.Rows.Count > 0 Then

            For i As Integer = 0 To dtVendor.Rows.Count - 1
                vData = New VendorFullDetail_Data
                vData.VendorID = dtVendor.Rows(i)("VendorID")
                vData.VendorGroupID = dtVendor.Rows(i)("VendorGroupID")
                vData.VendorCode = dtVendor.Rows(i)("VendorCode")
                vData.VendorName = dtVendor.Rows(i)("VendorName")
                vData.VendorGroupCode = dtVendor.Rows(i)("VendorGroupCode")
                vData.VendorGroupName = dtVendor.Rows(i)("VendorGroupName")
                If IsDBNull(dtVendor.Rows(i)("VendorFirstName")) Then
                    dtVendor.Rows(i)("VendorFirstName") = ""
                End If
                vData.VendorFirstName = dtVendor.Rows(i)("VendorFirstName")
                If IsDBNull(dtVendor.Rows(i)("VendorLastName")) Then
                    dtVendor.Rows(i)("VendorLastName") = ""
                End If
                vData.VendorLastName = dtVendor.Rows(i)("VendorLastName")
                If IsDBNull(dtVendor.Rows(i)("VendorAddress1")) Then
                    dtVendor.Rows(i)("VendorAddress1") = ""
                End If
                vData.VendorAddress1 = dtVendor.Rows(i)("VendorAddress1")
                If IsDBNull(dtVendor.Rows(i)("VendorAddress2")) Then
                    dtVendor.Rows(i)("VendorAddress2") = ""
                End If
                vData.VendorAddress2 = dtVendor.Rows(i)("VendorAddress2")
                If IsDBNull(dtVendor.Rows(i)("VendorCity")) Then
                    dtVendor.Rows(i)("VendorCity") = ""
                End If
                vData.VendorCity = dtVendor.Rows(i)("VendorCity")
                If IsDBNull(dtVendor.Rows(i)("ProvinceName")) Then
                    dtVendor.Rows(i)("ProvinceName") = ""
                End If
                If IsDBNull(dtVendor.Rows(i)("VendorProvince")) Then
                    dtVendor.Rows(i)("VendorProvince") = 0
                End If
                vData.VendorProvinceID = dtVendor.Rows(i)("VendorProvince")
                vData.VendorProvinceName = dtVendor.Rows(i)("ProvinceName")
                If IsDBNull(dtVendor.Rows(i)("VendorZipCode")) Then
                    dtVendor.Rows(i)("VendorZipCode") = ""
                End If
                vData.VendorZipCode = dtVendor.Rows(i)("VendorZipCode")
                If IsDBNull(dtVendor.Rows(i)("VendorTelephone")) Then
                    dtVendor.Rows(i)("VendorTelephone") = ""
                End If
                vData.VendorTelephone = dtVendor.Rows(i)("VendorTelephone")
                If IsDBNull(dtVendor.Rows(i)("VendorFax")) Then
                    dtVendor.Rows(i)("VendorFax") = ""
                End If
                vData.VendorFax = dtVendor.Rows(i)("VendorFax")
                If IsDBNull(dtVendor.Rows(i)("VendorMobile")) Then
                    dtVendor.Rows(i)("VendorMobile") = ""
                End If
                vData.VendorMobile = dtVendor.Rows(i)("VendorMobile")
                If IsDBNull(dtVendor.Rows(i)("VendorEmail")) Then
                    dtVendor.Rows(i)("VendorEmail") = ""
                End If
                vData.VendorEMail = dtVendor.Rows(i)("VendorEmail")
                If IsDBNull(dtVendor.Rows(i)("VendorAdditional")) Then
                    dtVendor.Rows(i)("VendorAdditional") = ""
                End If
                vData.VendorAddtional = dtVendor.Rows(i)("VendorAdditional")
                If IsDBNull(dtVendor.Rows(i)("VendorTermOfPayment")) Then
                    dtVendor.Rows(i)("VendorTermOfPayment") = 0
                End If
                vData.TermOfPayment = dtVendor.Rows(i)("VendorTermOfPayment")
                If IsDBNull(dtVendor.Rows(i)("VendorCreditDay")) Then
                    dtVendor.Rows(i)("VendorCreditDay") = 0
                End If
                vData.CreditDay = dtVendor.Rows(i)("VendorCreditDay")
                vData.DefaultTaxType = dtVendor.Rows(i)("MaterialTaxType")
                If Not IsDBNull(dtVendor.Rows(0)("MaterialTaxTypeName")) Then
                    vData.DefaultTaxTypeName = dtVendor.Rows(i)("MaterialTaxTypeName")
                End If
                vListData.Add(vData)
            Next
        End If

        Return vListData
    End Function
End Class