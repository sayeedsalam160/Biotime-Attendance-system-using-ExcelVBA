Attribute VB_Name = "modMain"
Public Sub GenerateAttendance()

    If Not SelectACSFile Then Exit Sub

    ImportACSData

    Call BuildHeaderMap

    If Not ValidateHeaders Then Exit Sub
    
    Call ProcessMorningShift

End Sub
