Module MaterialModule

    Friend Function SearchMaterialByCodeOrName(ByVal globalVariable As GlobalVariable, ByVal keyWord As String, ByVal isSearchCode As Boolean, ByRef materialList As List(Of ListMaterialDetail_Data), ByRef resultText As String) As Boolean
        Dim dtMaterials As New DataTable
        Dim dtUnits As New DataTable
        Dim dtMaterialDefaultPrice As New DataTable
        Dim defaultTaxType As Integer = 99
        dtUnits = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn)
        If isSearchCode = False Then
            dtMaterials = MaterialSQL.SearchMaterialByName(globalVariable.DocDBUtil, globalVariable.DocConn, keyWord)
        Else
            dtMaterials = MaterialSQL.SearchMaterialByCode(globalVariable.DocDBUtil, globalVariable.DocConn, keyWord)
        End If
        Return MaterialModule.InsertMaterialFromDataTableToList(globalVariable, dtMaterials, dtUnits, dtMaterialDefaultPrice, defaultTaxType, materialList, resultText)
    End Function

    Friend Function InsertMaterialFromDataTableToList(ByVal globalVariable As GlobalVariable, ByVal dtMaterial As DataTable, ByVal dtMaterialUnit As DataTable,
                                                      ByVal dtMaterialDefaultPrice As DataTable, ByVal defaultTaxType As Integer, ByRef materialList As List(Of ListMaterialDetail_Data),
                                                      ByRef resultText As String) As Boolean
        Dim i, j As Integer
        Dim expression As String = ""
        Dim foundRows() As DataRow
        Dim expressionPrice As String = ""
        Dim foundRowsPrice() As DataRow
        Dim defaultPrice As Decimal = 0
        Dim materialUnitList As List(Of ListMaterialUnit_Data)
        Dim defaultUnitLargeId As Integer = 0
        Dim defaultUnitLargeName As String = ""

        Try
            materialList = New List(Of ListMaterialDetail_Data)
            If dtMaterial.Rows.Count > 0 Then

                For i = 0 To dtMaterial.Rows.Count - 1
                    defaultUnitLargeName = ""
                    defaultUnitLargeId = 0
                    If IsDBNull(dtMaterial.Rows(i)("MaterialCode")) Then
                        dtMaterial.Rows(i)("MaterialCode") = ""
                    End If
                    If IsDBNull(dtMaterial.Rows(i)("MaterialName")) Then
                        dtMaterial.Rows(i)("MaterialName") = ""
                    End If

                    materialUnitList = New List(Of ListMaterialUnit_Data)
                    expression = "MaterialID=" & dtMaterial.Rows(i)("MaterialID")
                    foundRows = dtMaterialUnit.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        For j = 0 To foundRows.Length - 1
                            If dtMaterialDefaultPrice.Rows.Count > 0 Then
                                expressionPrice = "MaterialID=" & dtMaterial.Rows(i)("MaterialID") & " AND SelectUnitLargeID=" & foundRows(j)("UnitLargeID")
                                foundRowsPrice = dtMaterialDefaultPrice.Select(expressionPrice)
                                If foundRowsPrice.GetUpperBound(0) >= 0 Then
                                    defaultPrice = foundRowsPrice(0)("DefaultPrice")
                                Else
                                    defaultPrice = 0
                                End If
                            End If
                            materialUnitList.Add(ListMaterialUnit_Data.NewMaterialUnit(foundRows(j)("UnitSmallID"), foundRows(j)("UnitSmallName"),
                                                foundRows(j)("UnitSmallRatio"), foundRows(j)("UnitLargeID"), foundRows(j)("UnitLargeName"),
                                                foundRows(j)("UnitLargeRatio"), foundRows(j)("IsDefault"), defaultPrice))

                            
                        Next j
                        defaultUnitLargeId = dtMaterial.Rows(i)("SAPUnitID")
                        defaultUnitLargeName = foundRows(0)("UnitLargeName")
                       
                    End If
                    If defaultTaxType = 99 Then
                        defaultTaxType = dtMaterial.Rows(i)("MaterialTaxType")
                    End If
                    materialList.Add(ListMaterialDetail_Data.NewListMaterial(dtMaterial.Rows(i)("MaterialID"), dtMaterial.Rows(i)("MaterialDeptID"),
                                    dtMaterial.Rows(i)("MaterialCode"), dtMaterial.Rows(i)("MaterialName"), defaultTaxType,
                                    dtMaterial.Rows(i)("UnitSmallID"), defaultUnitLargeId, defaultUnitLargeName, dtMaterial.Rows(i)("SAPUnitID"), materialUnitList,
                                    ListMaterialTaxType_Data.ListMaterialTaxType))
                Next i
            Else
                resultText = GlobalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If

        Catch ex As Exception
            resultText = ex.ToString
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "MaterialModule", "InsertMaterialFromDataTableToList", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function InsertMaterialUnitFromDataTableToList(ByVal globalVariable As GlobalVariable, ByVal dtMaterialUnit As DataTable, ByVal dtMaterialDefaultPrice As DataTable,
                                                          ByRef materialUnitList As List(Of ListMaterialUnit_Data), ByRef resultText As String) As Boolean
        Dim j As Integer
        Dim expressionPrice As String = ""
        Dim foundRowsPrice() As DataRow
        Dim defaultPrice As Decimal = 0

        Try
            materialUnitList = New List(Of ListMaterialUnit_Data)
            If dtMaterialUnit.Rows.Count > 0 Then
                For j = 0 To dtMaterialUnit.Rows.Count - 1
                    If dtMaterialDefaultPrice.Rows.Count > 0 Then
                        expressionPrice = "MaterialID=" & dtMaterialUnit.Rows(j)("MaterialID") & " AND SelectUnitLargeID=" & dtMaterialUnit.Rows(j)("UnitLargeID")
                        foundRowsPrice = dtMaterialDefaultPrice.Select(expressionPrice)
                        If foundRowsPrice.GetUpperBound(0) >= 0 Then
                            defaultPrice = foundRowsPrice(0)("DefaultPrice")
                        Else
                            defaultPrice = 0
                        End If
                    End If
                    materialUnitList.Add(ListMaterialUnit_Data.NewMaterialUnit(dtMaterialUnit.Rows(j)("UnitSmallID"), dtMaterialUnit.Rows(j)("UnitSmallName"),
                                        dtMaterialUnit.Rows(j)("UnitSmallRatio"), dtMaterialUnit.Rows(j)("UnitLargeID"), dtMaterialUnit.Rows(j)("UnitLargeName"),
                                        dtMaterialUnit.Rows(j)("UnitLargeRatio"), dtMaterialUnit.Rows(j)("IsDefault"), defaultPrice))
                Next j
            Else
                resultText = GlobalVariable.MESSAGE_DATANOTFOUND
                Return False
            End If
        Catch ex As Exception
            resultText = ex.ToString
            DocumentSQL.InsertLog(globalVariable.DocDBUtil, globalVariable.DocConn, "MaterialModule", "InsertMaterialUnitFromDataTableToList", "99", ex.ToString)
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function MaterialAmountInLargeUnit(ByVal dtDetail As DataTable, ByVal materialID As Integer,   ByVal materialAmount As Decimal, ByVal unitSmallName As String) As String
        Dim rResult() As DataRow
        Dim dclAmount As Decimal
        Dim dclCal, dclRatio As Decimal
        Dim i As Integer
        Dim strTemp As String
        If materialAmount = 0 Then
            Return Format(materialAmount, "#,##0.##") & " " & unitSmallName
        End If
        'Calculate Material In eahc unit
        If materialAmount < 0 Then
            dclAmount = (materialAmount * -1)
        Else
            dclAmount = materialAmount
        End If
        strTemp = ""
        rResult = dtDetail.Select("MaterialID = " & materialID)
        For i = 0 To rResult.Length - 1
            dclRatio = rResult(i)("UnitSmallRatio") / rResult(i)("UnitLargeRatio")
            If dclRatio = 0 Then
                dclRatio = 1
            End If
            If dclRatio <> 1 Then
                dclCal = dclAmount \ dclRatio
                dclAmount = dclAmount Mod dclRatio
                If dclCal > 0 Then
                    strTemp &= Format(dclCal, "#,##0") & " " & rResult(i)("UnitLargeName") & " "
                End If
            Else
                strTemp &= Format(dclAmount, "#,##0") & " " & rResult(i)("UnitLargeName") & " "
                dclAmount = 0
            End If
            'Amount left is 0, no need to convert the other
            If dclAmount = 0 Then
                If materialAmount < 0 Then
                    strTemp = "-" & strTemp
                End If
                Return Trim(strTemp)
            End If
        Next i
        If dclAmount <> 0 Then
            strTemp &= Format(dclAmount, "#,##0.##") & " " & unitSmallName
        End If
        If materialAmount < 0 Then
            strTemp = "-" & strTemp
        End If
        Return Trim(strTemp)
    End Function
End Module
