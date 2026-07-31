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

    ProcessMorningShift
    WriteMorningShiftDatabase

    ProcessNightShift
    WriteNightShiftDatabase

End Sub
