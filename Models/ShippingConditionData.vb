Public Class ShippingCondition_Data
    Public ShippingCondition As String

    Public Shared Function NewShippingCondition(ByVal shippingCondition As String) As ShippingCondition_Data
        Dim data As New ShippingCondition_Data
        data.ShippingCondition = shippingCondition
        Return data
    End Function
End Class
