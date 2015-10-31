Public Class ListMaterialDept_Data
    Public MaterialDeptID As Integer
    Public MaterialGroupID As Integer
    Public MaterialDeptCode As String
    Public MaterialDeptName As String

    Public Shared Function NewListMaterialDept(ByVal deptID As Integer, ByVal groupID As Integer,
    ByVal deptCode As String, ByVal deptName As String) As ListMaterialDept_Data
        Dim mData As New ListMaterialDept_Data
        mData.MaterialDeptID = deptID
        mData.MaterialGroupID = groupID
        mData.MaterialDeptCode = deptCode
        mData.MaterialDeptName = deptName
        Return mData
    End Function
End Class