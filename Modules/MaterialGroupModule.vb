Imports System
Imports System.Data
Imports pRoMiSe.Utilitys.Utilitys

Module MaterialGroupModule

    Friend Function ListMaterialGroup(ByVal globalVariable As GlobalVariable, ByRef materialGroupList As List(Of ListMaterialGroup_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtMaterialGroup As DataTable

        Try
            dtMaterialGroup = MaterialSQL.ListMaterialGroup(globalVariable.DocDBUtil, globalVariable.DocConn)
            materialGroupList = New List(Of ListMaterialGroup_Data)
            If dtMaterialGroup.Rows.Count > 0 Then
                For i = 0 To dtMaterialGroup.Rows.Count - 1
                    If IsDBNull(dtMaterialGroup.Rows(i)("MaterialGroupCode")) Then
                        dtMaterialGroup.Rows(i)("MaterialGroupCode") = ""
                    End If
                    If IsDBNull(dtMaterialGroup.Rows(i)("MaterialGroupName")) Then
                        dtMaterialGroup.Rows(i)("MaterialGroupName") = ""
                    End If
                    materialGroupList.Add(ListMaterialGroup_Data.NewListMaterialGroup(dtMaterialGroup.Rows(i)("MaterialGroupID"),
                                          dtMaterialGroup.Rows(i)("MaterialGroupType"), dtMaterialGroup.Rows(i)("MaterialGroupCode"),
                                          dtMaterialGroup.Rows(i)("MaterialGroupName")))
                Next i
            Else
                resultText = GlobalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.ToString
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "MaterialGroupModule", "ListMaterialGroup", "99", ex.ToString)
            Return False
        End Try

        resultText = ""
        Return True
    End Function
   
End Module
