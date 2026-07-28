Attribute VB_Name = "modMorningShift"
Option Explicit

Public Sub ProcessMorningShift()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim RowNum As Long
    
    Dim EmployeeID As String
    Dim EmployeeName As String
    Dim PunchDate As Variant
    Dim PunchTime As Variant
    Dim PunchState As String

    Set ws = ThisWorkbook.Worksheets("Raw Import")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For RowNum = DATA_START_ROW To LastRow

        EmployeeID = Trim(ws.Cells(RowNum, gHeaderMap("Employee ID")).Value)
        EmployeeName = Trim(ws.Cells(RowNum, gHeaderMap("First Name")).Value)
        PunchDate = ws.Cells(RowNum, gHeaderMap("Date")).Value
        PunchTime = ws.Cells(RowNum, gHeaderMap("Time")).Value
        PunchState = Trim(ws.Cells(RowNum, gHeaderMap("Punch State")).Value)

        If Not IsValidEmployeeID(EmployeeID) Then GoTo NextRecord
        If Not IsValidDateField(PunchDate) Then GoTo NextRecord
        If Not IsValidTimeField(PunchTime) Then GoTo NextRecord

NextRecord:
    Next RowNum

End Sub
