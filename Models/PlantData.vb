Public Class Plant_Data
    Public PlantID As Integer
    Public PlantName As String

    Public Shared Function NewPlant(ByVal plantId As Integer, ByVal plantName As String) As Plant_Data
        Dim data As New Plant_Data
        data.PlantID = plantId
        data.PlantName = plantName
        Return data
    End Function
End Class