Public Class BusinessPlace_Data
    Public BusinessPlaceID As Integer
    Public BusinessPlaceName As String

    Public Shared Function NewBusinessPlace(ByVal businessPlaceId As Integer, ByVal businessPlaceName As String) As BusinessPlace_Data
        Dim data As New BusinessPlace_Data
        data.BusinessPlaceID = businessPlaceId
        data.BusinessPlaceName = businessPlaceId & " " & businessPlaceName
        Return data
    End Function
End Class