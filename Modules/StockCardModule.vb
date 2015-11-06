Imports pRoMiSe.Utilitys.Utilitys

Module StockCardModule

    Friend Function GetStockCardReport(ByVal globalVariable As GlobalVariable, ByVal inventoryId As Integer, ByVal materialGroupId As String, ByVal materialDeptId As String, ByVal materialCode As String, ByVal startDate As Date, ByVal endDate As Date, ByRef stockCardData As List(Of StockCard_Data), ByRef resultText As String) As Boolean


        Dim strStartDate, strEndDate As String
        Dim dtTable As New DataTable
        Dim dtUnit As New DataTable
        Dim dummyTable As String = ""
        Dim stockData As New List(Of StockCard_Data)
        Dim data As New StockCard_Data
        Dim selYear, selMonth As Integer
        Dim costTypeVal As Integer = 1
        Dim expression As String = ""
        Dim foundRows() As DataRow
        Dim UnitRatioVal As Double = 1
        Dim CostPerUnit As Double
        Dim groupData As New DataTable
        Dim GroupDataL As New DataTable
        Dim OnhandStock As Decimal = 0
        Dim varianceStock As Decimal = 0
        Dim checkCountStock As Boolean = False
        dummyTable = "dummy" & GenerateGUID().Replace("-", "")
        strStartDate = FormatDate(startDate)
        strEndDate = FormatDate(endDate)

        selYear = startDate.ToString("yyyy", globalVariable.InvariantCulture)
        selMonth = startDate.ToString("MM", globalVariable.InvariantCulture)

        Try
            DocumentSQL.CreateDummyStockCard(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryId, startDate, endDate)
            DocumentSQL.CrateDummyMaterialStdCost(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryId, selMonth, selYear, costTypeVal, dummyTable)
            dtTable = DocumentSQL.GetMaterialStock(globalVariable.DocDBUtil, globalVariable.DocConn, materialGroupId, materialDeptId, materialCode, dummyTable, groupData, GroupDataL)
            DocumentSQL.DropTempTable(globalVariable.DocDBUtil, globalVariable.DocConn, dummyTable)
            dtUnit = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn)
            checkCountStock = DocumentSQL.CheckCountStock(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryId, strEndDate)
            If dtTable.Rows.Count > 0 Then
                For i As Integer = 0 To dtTable.Rows.Count - 1
                    data = New StockCard_Data
                    OnhandStock = 0
                    varianceStock = 0
                    UnitRatioVal = 1
                    If Not IsDBNull(dtTable.Rows(i)("MaterialGroupId")) Then
                        data.MaterialGroupId = dtTable.Rows(i)("MaterialGroupId")
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("MaterialGroupCode")) Then
                        data.MaterialGroupCode = dtTable.Rows(i)("MaterialGroupCode")
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("MaterialGroupName")) Then
                        data.MaterialGroupName = dtTable.Rows(i)("MaterialGroupName")
                    End If
                    data.MaterialId = dtTable.Rows(i)("MaterialID")
                    data.MaterialCode = dtTable.Rows(i)("MaterialCode")
                    data.MaterialName = dtTable.Rows(i)("MaterialName")
                    expression = "MaterialId=" & dtTable.Rows(i)("MaterialId") & " And UnitSmallId=" & dtTable.Rows(i)("UnitSmallId")
                    foundRows = dtUnit.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        UnitRatioVal = foundRows(0)("UnitSmallRatio")
                        If Not IsDBNull(foundRows(0)("UnitLargeName")) Then
                            data.UnitName = foundRows(0)("UnitLargeName")
                        End If
                    End If
                    CostPerUnit = 0
                    If Not IsDBNull(dtTable.Rows(i)("TotalPrice")) And Not IsDBNull(dtTable.Rows(i)("TotalAmount")) Then
                        If dtTable.Rows(i)("TotalAmount") > 0 Then
                            CostPerUnit = dtTable.Rows(i)("TotalPrice") / dtTable.Rows(i)("TotalAmount")
                        End If
                    End If
                    data.UnitCost = Format(CostPerUnit * UnitRatioVal, "##,##0.00")
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount0")) Then
                        data.BeginningAmount = dtTable.Rows(i)("NetSmallAmount0")
                        OnhandStock += data.BeginningAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount1")) Then
                        data.ReceiveAmount = Format(dtTable.Rows(i)("NetSmallAmount1") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.ReceiveAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount2")) Then
                        data.TransferInAmount = Format(dtTable.Rows(i)("NetSmallAmount2") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.TransferInAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount3")) Then
                        data.TransferOutAmount = Format(dtTable.Rows(i)("NetSmallAmount3") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.TransferOutAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount4")) Then
                        data.AdjustAmount = Format(dtTable.Rows(i)("NetSmallAmount4") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.AdjustAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount5")) Then
                        data.DestroyAmount = Format(dtTable.Rows(i)("NetSmallAmount5") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.DestroyAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount6")) Then
                        data.SaleAmount = Format(dtTable.Rows(i)("NetSmallAmount6") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.SaleAmount
                    End If
                    data.OnhandAmount = OnhandStock
                    If checkCountStock = True Then
                        If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount7")) Then
                            data.VarianceAmount = Format(dtTable.Rows(i)("NetSmallAmount7") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                            varianceStock = data.VarianceAmount
                        End If
                        data.EndingAmount = (data.OnhandAmount + varianceStock)
                    End If
                    stockCardData.Add(data)
                Next
            End If

        Catch ex As Exception
            resultText = ex.ToString
            Return False
        End Try
        resultText = ""
        Return True
    End Function

    Friend Function GetStockCardReportDetail(ByVal globalVariable As GlobalVariable, ByVal inventoryId As Integer, ByVal materialGroupId As String,
                                       ByVal materialDeptId As String, ByVal materialCode As String, ByVal startDate As Date,
                                       ByVal endDate As Date, ByRef stockCardData As List(Of StockCard_Data),
                                       ByRef resultText As String) As Boolean


        Dim strStartDate, strEndDate As String
        Dim dtTable As New DataTable
        Dim dtUnit As New DataTable
        Dim dummyTable As String = ""
        Dim stockData As New List(Of StockCard_Data)
        Dim data As New StockCard_Data
        Dim selYear, selMonth As Integer
        Dim costTypeVal As Integer = 1
        Dim expression As String = ""
        Dim foundRows() As DataRow
        Dim UnitRatioVal As Double = 1
        Dim CostPerUnit As Double
        Dim groupData As New DataTable
        Dim GroupDataL As New DataTable
        Dim OnhandStock As Decimal = 0
        Dim varianceStock As Decimal = 0

        dummyTable = "dummy" & GenerateGUID().Replace("-", "")
        strStartDate = FormatDate(startDate)
        strEndDate = FormatDate(endDate)

        selYear = startDate.ToString("yyyy", globalVariable.InvariantCulture)
        selMonth = startDate.ToString("MM", globalVariable.InvariantCulture)

        Try
            DocumentSQL.CreateDummyStockCard(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryId, startDate, endDate)
            DocumentSQL.CrateDummyMaterialStdCost(globalVariable.DocDBUtil, globalVariable.DocConn, inventoryId, selMonth, selYear, costTypeVal, dummyTable)
            dtTable = DocumentSQL.GetMaterialStock(globalVariable.DocDBUtil, globalVariable.DocConn, materialGroupId, materialDeptId, materialCode, dummyTable, groupData, GroupDataL)
            DocumentSQL.DropTempTable(globalVariable.DocDBUtil, globalVariable.DocConn, dummyTable)
            dtUnit = MaterialSQL.ListMaterialUnit(globalVariable.DocDBUtil, globalVariable.DocConn)

            If dtTable.Rows.Count > 0 Then
                For i As Integer = 0 To dtTable.Rows.Count - 1
                    data = New StockCard_Data
                    OnhandStock = 0
                    UnitRatioVal = 1
                    If Not IsDBNull(dtTable.Rows(i)("MaterialGroupId")) Then
                        data.MaterialGroupId = dtTable.Rows(i)("MaterialGroupId")
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("MaterialGroupCode")) Then
                        data.MaterialGroupCode = dtTable.Rows(i)("MaterialGroupCode")
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("MaterialGroupName")) Then
                        data.MaterialGroupName = dtTable.Rows(i)("MaterialGroupName")
                    End If
                    data.MaterialId = dtTable.Rows(i)("MaterialID")
                    data.MaterialCode = dtTable.Rows(i)("MaterialCode")
                    data.MaterialName = dtTable.Rows(i)("MaterialName")
                    expression = "MaterialId=" & dtTable.Rows(i)("MaterialId") & " And UnitSmallId=" & dtTable.Rows(i)("UnitSmallId")
                    foundRows = dtUnit.Select(expression)
                    If foundRows.GetUpperBound(0) >= 0 Then
                        UnitRatioVal = foundRows(0)("UnitSmallRatio")
                        If Not IsDBNull(foundRows(0)("UnitLargeName")) Then
                            data.UnitName = foundRows(0)("UnitLargeName")
                        End If
                    End If
                    CostPerUnit = 0
                    If Not IsDBNull(dtTable.Rows(i)("TotalPrice")) And Not IsDBNull(dtTable.Rows(i)("TotalAmount")) Then
                        If dtTable.Rows(i)("TotalAmount") > 0 Then
                            CostPerUnit = dtTable.Rows(i)("TotalPrice") / dtTable.Rows(i)("TotalAmount")
                        End If
                    End If
                    data.UnitCost = Format(CostPerUnit * UnitRatioVal, "##,##0.00")
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount0")) Then
                        data.BeginningAmount = dtTable.Rows(i)("NetSmallAmount0")
                        OnhandStock += data.BeginningAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount1")) Then
                        data.ReceiveAmount = Format(dtTable.Rows(i)("NetSmallAmount1") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.ReceiveAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount2")) Then
                        data.TransferInAmount = Format(dtTable.Rows(i)("NetSmallAmount2") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.TransferInAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount3")) Then
                        data.TransferOutAmount = Format(dtTable.Rows(i)("NetSmallAmount3") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.TransferOutAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount4")) Then
                        data.AdjustAmount = Format(dtTable.Rows(i)("NetSmallAmount4") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.AdjustAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount5")) Then
                        data.DestroyAmount = Format(dtTable.Rows(i)("NetSmallAmount5") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.DestroyAmount
                    End If
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount6")) Then
                        data.SaleAmount = Format(dtTable.Rows(i)("NetSmallAmount6") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        OnhandStock += data.SaleAmount
                    End If
                    data.OnhandAmount = OnhandStock
                    If Not IsDBNull(dtTable.Rows(i)("NetSmallAmount7")) Then
                        data.VarianceAmount = Format(dtTable.Rows(i)("NetSmallAmount7") / UnitRatioVal, "##,##0.00;(##,##0.00)")
                        varianceStock = data.VarianceAmount
                    End If
                    If varianceStock <> 0 Then
                        data.EndingAmount = (data.OnhandAmount + varianceStock)
                    End If

                    stockCardData.Add(data)
                Next
            End If

        Catch ex As Exception
            resultText = ex.ToString
            Return False
        End Try
        resultText = ""
        Return True
    End Function

End Module
