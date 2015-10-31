Public Class Customer_Data
    Public CustomerCode As String
    Public CustomerName As String
    Public CustomerAddress As String
    Public CustomerShipTo As String
    Public CustomerBillTo As String

    Public Shared Function NewCustomerData(ByVal customerCode As String, ByVal customerName As String, ByVal customerAddress As String,
    ByVal customerShipTo As String, ByVal customerBillTo As String) As Customer_Data

        Dim data As New Customer_Data
        data.CustomerCode = customerCode
        data.CustomerName = customerName
        data.CustomerAddress = customerAddress
        data.CustomerShipTo = customerShipTo
        data.CustomerBillTo = customerBillTo
        Return data
    End Function
End Class
