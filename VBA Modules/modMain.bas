Attribute VB_Name = "modMain"
Option Explicit

Public Sub GenerateAttendance()

    If Not SelectACSFile Then Exit Sub

    ImportACSData

End Sub
