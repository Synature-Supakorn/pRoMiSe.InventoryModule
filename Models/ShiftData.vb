Public Class Shift_Data
    Public ShiftID As Integer
    Public ShiftNo As String

    Public Shared Function NewShift(ByVal shiftId As Integer, ByVal shiftNo As String) As Shift_Data
        Dim data As New Shift_Data
        data.ShiftID = shiftId
        data.ShiftNo = shiftNo
        Return data
    End Function
End Class

