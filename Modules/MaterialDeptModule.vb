Module MaterialDeptModule

    Friend Function ListMaterialDept(ByVal globalVariable As GlobalVariable, ByVal materialGroupID As Integer, ByRef materialDeptList As List(Of ListMaterialDept_Data), ByRef resultText As String) As Boolean
        Dim i As Integer
        Dim dtMaterialDept As DataTable
        Try
            dtMaterialDept = MaterialSQL.ListMaterialDept(globalVariable.DocDBUtil, globalVariable.DocConn, materialGroupID)
            materialDeptList = New List(Of ListMaterialDept_Data)
            If dtMaterialDept.Rows.Count > 0 Then
                For i = 0 To dtMaterialDept.Rows.Count - 1
                    If IsDBNull(dtMaterialDept.Rows(i)("MaterialDeptCode")) Then
                        dtMaterialDept.Rows(i)("MaterialDeptCode") = ""
                    End If
                    If IsDBNull(dtMaterialDept.Rows(i)("MaterialDeptName")) Then
                        dtMaterialDept.Rows(i)("MaterialDeptName") = ""
                    End If
                    materialDeptList.Add(ListMaterialDept_Data.NewListMaterialDept(dtMaterialDept.Rows(i)("MaterialDeptID"),
                                         dtMaterialDept.Rows(i)("MaterialGroupID"), dtMaterialDept.Rows(i)("MaterialDeptCode"),
                                         dtMaterialDept.Rows(i)("MaterialDeptName")))
                Next i
            Else
                resultText = GlobalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.ToString
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "ListMaterialDept", "ListMaterialDept", "99", ex.ToString)
            Return False
        End Try

        resultText = ""
        Return True
    End Function

End Module
