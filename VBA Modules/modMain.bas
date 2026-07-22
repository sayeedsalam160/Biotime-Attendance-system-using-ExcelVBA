Attribute VB_Name = "modMain"
Option Explicit

Public Sub GenerateAttendance()

    If Not SelectACSFile Then Exit Sub

    MsgBox "Selected File:" & vbCrLf & _
           gACSFilePath, _
           vbInformation, _
           "Attendance Engine"

End Sub
