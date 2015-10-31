Public Class ListMaterialGroup_Data

    Public MaterialGroupID As Integer
    Public MaterialGroupType As Integer
    Public MaterialGroupCode As String
    Public MaterialGroupName As String

    Public Shared Function NewListMaterialGroup(ByVal groupID As Integer, ByVal groupType As Integer,
                                                ByVal groupCode As String, ByVal groupName As String) As ListMaterialGroup_Data
        Dim mData As New ListMaterialGroup_Data
        mData.MaterialGroupID = groupID
        mData.MaterialGroupType = groupType
        mData.MaterialGroupCode = groupCode
        mData.MaterialGroupName = groupName
        Return mData
    End Function
End Class