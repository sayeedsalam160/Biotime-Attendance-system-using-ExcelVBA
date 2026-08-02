Attribute VB_Name = "modMain"
Option Explicit

Public Sub GenerateAttendance()

    If Not SelectACSFile Then Exit Sub

    ImportACSData

    BuildHeaderMap

    If Not ValidateHeaders Then Exit Sub

    UpdateRawPunchDatabase

    ClearAttendanceDatabase "Morning Shift Database"
    ClearAttendanceDatabase "Night Shift Database"

    'ProcessMorningShift
    'WriteMorningShiftDatabase

    'ProcessNightShift
    'WriteNightShiftDatabase
    
    ProcessAttendanceEngine
    
        UpdateTodaysSummary

    MsgBox _
        "Attendance processing completed successfully." & vbCrLf & _
        "The attendance databases and management summary have been updated.", _
        vbInformation, _
        "Attendance Automation"

End Sub
