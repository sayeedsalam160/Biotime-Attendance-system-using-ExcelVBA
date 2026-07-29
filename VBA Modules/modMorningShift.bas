Attribute VB_Name = "modMorningShift"
Option Explicit

Public Sub ProcessMorningShift()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim RowNum As Long

    Dim EmployeeID As String
    Dim PunchDate As Date
    Dim PunchTime As Date
    Dim AttendanceDate As Date
    Dim DictKey As String

    Dim PunchData As Variant
    Dim Key As Variant
    Dim Counter As Long

    Set ws = ThisWorkbook.Worksheets("Raw Import")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set gMorningShiftDict = CreateObject("Scripting.Dictionary")

    For RowNum = DATA_START_ROW To LastRow

        EmployeeID = Trim(ws.Cells(RowNum, gHeaderMap("Employee ID")).Value)

        PunchDate = ws.Cells(RowNum, gHeaderMap("Date")).Value

        PunchTime = ws.Cells(RowNum, gHeaderMap("Time")).Value

        AttendanceDate = GetAttendanceDate(PunchDate, PunchTime)

        DictKey = EmployeeID & "|" & Format(AttendanceDate, "dd-mm-yyyy")
        
        If PunchTime < TimeSerial(12, 0, 0) Then

            Debug.Print EmployeeID & _
                " | " & PunchDate & _
                " | " & PunchTime & _
                " | Attendance=" & AttendanceDate

        End If

        'Debug first 10 records only
        If RowNum <= 10 Then

            Debug.Print "PunchDate      = " & PunchDate
            Debug.Print "PunchTime      = " & PunchTime
            Debug.Print "AttendanceDate = " & AttendanceDate
            Debug.Print "DictKey        = " & DictKey
            Debug.Print "-------------------------"

        End If

        If Not gMorningShiftDict.Exists(DictKey) Then

            gMorningShiftDict.Add DictKey, Array(PunchTime, PunchTime)

        Else

            PunchData = gMorningShiftDict(DictKey)

            If PunchTime < PunchData(0) Then
                PunchData(0) = PunchTime
            End If

            If PunchTime > PunchData(1) Then
                PunchData(1) = PunchTime
            End If

            gMorningShiftDict(DictKey) = PunchData

        End If

    Next RowNum

    Debug.Print "Last Row = " & LastRow
    Debug.Print "Dictionary Count = " & gMorningShiftDict.Count
    Debug.Print "=============================="

    Counter = 0

    For Each Key In gMorningShiftDict.Keys

        PunchData = gMorningShiftDict(Key)

        Debug.Print Key & _
                    " | First=" & PunchData(0) & _
                    " | Last=" & PunchData(1)

        Counter = Counter + 1

        If Counter >= 10 Then Exit For

    Next Key

    MsgBox gMorningShiftDict.Count & _
           " unique Employee-Date combinations found.", vbInformation

End Sub

