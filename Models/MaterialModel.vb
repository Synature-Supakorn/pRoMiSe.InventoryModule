Imports pRoMiSe.DBHelper
Imports pRoMiSe.Inventory

Public Class ListMaterialDetail_Data
    Public MaterialID As Integer
    Public MaterialDeptID As Integer
    Public MaterialCode As String
    Public MaterialName As String
    Public MaterialCode1 As String
    Public MaterialName1 As String
    Public MaterialTaxType As Integer
    Public ListMaterialTaxType As List(Of ListMaterialTaxType_Data)
    Public UnitSmallID As Integer
    Public AddUnitLargeID As Integer
    Public AddUnitLargeName As String
    Public SAPUnitID As Integer
    Public ListUnit As List(Of ListMaterialUnit_Data)

    Public Overloads Shared Function NewListMaterial(ByVal materialID As Integer, ByVal deptID As Integer, ByVal materialCode As String,
                                                     ByVal materialName As String, ByVal taxType As Integer, ByVal unitSmallID As Integer,
                                                     ByVal addUnitLargeID As Integer, ByVal addUnitLargeName As String, ByVal sapUnitId As Integer,
                                                     ListUnit As List(Of ListMaterialUnit_Data), ByVal materialTaxTypeList As List(Of ListMaterialTaxType_Data)) As ListMaterialDetail_Data
        Dim mData As New ListMaterialDetail_Data
        mData.MaterialID = materialID
        mData.MaterialDeptID = deptID
        mData.MaterialCode = materialCode
        mData.MaterialName = materialName
        mData.MaterialCode1 = materialCode
        mData.MaterialName1 = materialName
        mData.MaterialTaxType = taxType
        mData.UnitSmallID = unitSmallID
        mData.AddUnitLargeID = addUnitLargeID
        mData.AddUnitLargeName = addUnitLargeName
        mData.ListUnit = ListUnit
        mData.ListMaterialTaxType = materialTaxTypeList
        Return mData
    End Function
   
End Class

Public Class ListMaterialTaxType_Data
    Public MaterialTaxType As Integer
    Public MaterialTaxTypeName As String
    Private globalVariable As New GlobalVariable

    Shared Function NewListTaxType(ByVal materialTaxType As Integer, ByVal materialTaxTypeName As String) As ListMaterialTaxType_Data
        Dim mData As New ListMaterialTaxType_Data
        mData.MaterialTaxType = materialTaxType
        mData.MaterialTaxTypeName = materialTaxTypeName
        Return mData
    End Function

    Public Shared Function ListMaterialTaxType() As List(Of ListMaterialTaxType_Data)
        Dim taxtypeList As New List(Of ListMaterialTaxType_Data)
        taxtypeList.Add(ListMaterialTaxType_Data.NewListTaxType(1, "Include VAT"))
        taxtypeList.Add(ListMaterialTaxType_Data.NewListTaxType(2, "Exclude VAT"))
        taxtypeList.Add(ListMaterialTaxType_Data.NewListTaxType(0, "Non VAT"))
        Return taxtypeList
    End Function

    Public Shared Function ListMaterialTaxType(ByVal dt As DataTable) As List(Of ListMaterialTaxType_Data)
        Dim taxtypeList As New List(Of ListMaterialTaxType_Data)
        For i As Integer = 0 To dt.Rows.Count - 1
            taxtypeList.Add(ListMaterialTaxType_Data.NewListTaxType(dt.Rows(i)("MaterialTaxtype"), dt.Rows(i)("MaterialTaxTypeName")))
        Next
        Return taxtypeList
    End Function

End Class

Public Class ListMaterialUnit_Data
    Public UnitSmallID As Integer
    Public UnitSmallName As String
    Public UnitSmallRatio As Decimal
    Public UnitLargeID As Integer
    Public UnitLargeName As String
    Public UnitLargeRatio As Decimal
    Public DefaultPrice As Decimal
    Public IsDefault As Boolean

    Public Shared Function NewMaterialUnit(ByVal unitSmallID As Integer, ByVal unitSmallName As String, ByVal unitSmallRatio As Decimal, ByVal unitLargeID As Integer,
                                           ByVal unitLargeName As String, ByVal unitLargeRatio As Decimal, ByVal isDefault As Integer, ByVal defaultPrice As Decimal) As ListMaterialUnit_Data

        Dim mData As New ListMaterialUnit_Data
        mData.UnitSmallID = unitSmallID
        mData.UnitSmallName = unitSmallName
        mData.UnitSmallRatio = unitSmallRatio
        mData.UnitLargeID = unitLargeID
        mData.UnitLargeName = unitLargeName
        mData.UnitLargeRatio = unitLargeRatio
        If isDefault = 1 Then
            mData.IsDefault = True
        Else
            mData.IsDefault = False
        End If
        mData.DefaultPrice = defaultPrice
        Return mData
    End Function

End Class

Public Class ListMaterialDiscountType_Data
    Public MaterialDiscountType As Integer
    Public MaterialDiscountTypeName As String

    Public Shared Function NewListDiscountType(ByVal materialDiscountType As Integer,
                                          ByVal materialDiscountTypeName As String) As ListMaterialDiscountType_Data
        Dim mData As New ListMaterialDiscountType_Data
        mData.MaterialDiscountType = materialDiscountType
        mData.MaterialDiscountTypeName = materialDiscountTypeName
        Return mData
    End Function

    Public Shared Function ListMaterialDiscountType() As List(Of ListMaterialDiscountType_Data)
        Dim disCList As New List(Of ListMaterialDiscountType_Data)
        disCList.Add(ListMaterialDiscountType_Data.NewListDiscountType(1, "Baht"))
        disCList.Add(ListMaterialDiscountType_Data.NewListDiscountType(2, "%"))
        Return disCList
    End Function

End Class
 
Public Class MaterialNotEnoughStock_Data
    Public MaterialID As Integer
    Public MaterialCode As String
    Public MaterialName As String
    Public TransferSmallAmount As Decimal
    Public TransferDisplayText As String
    Public CurrentStockSmallAmount As Decimal
    Public CurrentStockDisplayText As String
    Public UnitSmallID As Integer

    Public Shared Function NewMaterialNotEnoughStock(ByVal materialID As Integer, ByVal materialCode As String, ByVal materialName As String,
                                                     ByVal transferSmallAmount As Decimal, ByVal transferDisplayText As String,
                                                     ByVal currentStockSmallAmount As Decimal, ByVal currentStockDisplayText As String,
                                                     ByVal unitSmallID As Integer) As MaterialNotEnoughStock_Data
        Dim mData As New MaterialNotEnoughStock_Data
        mData.MaterialID = materialID
        mData.MaterialCode = materialCode
        mData.MaterialName = materialName
        mData.TransferSmallAmount = transferSmallAmount
        mData.TransferDisplayText = transferDisplayText
        mData.CurrentStockSmallAmount = currentStockSmallAmount
        mData.CurrentStockDisplayText = currentStockDisplayText
        mData.UnitSmallID = unitSmallID
        Return mData
    End Function

End Class
