Public Class ListVendorGroup_Data
    Public VendorGroupID As Integer
    Public VendorShopID As Integer
    Public VendorGroupCode As String
    Public VendorGroupName As String

    Public Shared Function NewListVendorGroup(ByVal vendorGroupID As Integer, ByVal groupCode As String,
    ByVal groupName As String) As ListVendorGroup_Data
        Dim vData As New ListVendorGroup_Data
        vData.VendorGroupID = vendorGroupID
        vData.VendorShopID = 0
        vData.VendorGroupCode = groupCode
        vData.VendorGroupName = groupName
        Return vData
    End Function
End Class

