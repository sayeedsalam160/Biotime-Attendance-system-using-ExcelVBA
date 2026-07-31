Attribute VB_Name = "modAttendanceEngine"
Option Explicit

'=========================================================
' MODULE: modAttendanceEngine
'
' PURPOSE:
' Reads all historical punches from Raw Punch Database,
' builds employee attendance sessions, determines whether
' each session belongs to Morning or Night Shift, and writes
' the result to the correct attendance database.
'
' The engine supports:
' - Daily ACS reports
' - Weekly ACS reports
' - Monthly ACS reports
' - Repeated report imports
' - Previous-day punch IN with next-day punch OUT
'=========================================================

Private gEmployeePunches As Object
Private gEmployeeNames As Object
Private gSessionKeys As Object

'=========================================================
' MAIN ATTENDANCE ENGINE
'=========================================================
Public Sub ProcessAttendanceEngine()

    BuildEmployeePunchHistory
    BuildSessionKeys
    ProcessAllAttendanceSessions

End Sub

'=========================================================
' BUILD EMPLOYEE PUNCH HISTORY
'
' Dictionary structure:
'
' Employee ID
'     ?
' Collection of all FullDateTime punches
'=========================================================
Private Sub BuildEmployeePunchHistory()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim RowNum As Long

    Dim EmployeeID As String
    Dim EmployeeName As String

    Dim RawDate As Variant
    Dim RawTime As Variant

    Dim PunchDate As Date
    Dim PunchTime As Date
    Dim FullDateTime As Date

    Dim PunchCollection As Collection

    Set ws = ThisWorkbook.Worksheets("Raw Punch Database")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set gEmployeePunches = CreateObject("Scripting.Dictionary")
    Set gEmployeeNames = CreateObject("Scripting.Dictionary")

    For RowNum = DATA_START_ROW To LastRow

        EmployeeID = Trim$(CStr(ws.Cells(RowNum, 1).Value))
        EmployeeName = Trim$(CStr(ws.Cells(RowNum, 2).Value))

        RawDate = ws.Cells(RowNum, 4).Value
        RawTime = ws.Cells(RowNum, 5).Value

        If EmployeeID <> vbNullString Then

            If IsDate(RawDate) Then

                If IsDate(RawTime) Or IsNumeric(RawTime) Then

                    PunchDate = DateValue(CDate(RawDate))
                    PunchTime = TimeValue(CDate(RawTime))

                    FullDateTime = PunchDate + PunchTime

                    If Not gEmployeePunches.Exists(EmployeeID) Then

                        Set PunchCollection = New Collection

                        gEmployeePunches.Add _
                            EmployeeID, _
                            PunchCollection

                    End If

                    gEmployeePunches(EmployeeID).Add FullDateTime

                    If Not gEmployeeNames.Exists(EmployeeID) Then

                        gEmployeeNames.Add _
                            EmployeeID, _
                            EmployeeName

                    ElseIf gEmployeeNames(EmployeeID) = vbNullString Then

                        gEmployeeNames(EmployeeID) = EmployeeName

                    End If

                End If

            End If

        End If

    Next RowNum

End Sub

'=========================================================
' BUILD POSSIBLE ATTENDANCE SESSION KEYS
'
' For every punch date, two possible attendance dates
' are considered:
'
' 1. The punch date itself
' 2. The previous date
'
' This allows the engine to connect:
'
' Previous-day afternoon IN
'         +
' Next-day midnight OUT
'
' and:
'
' Next-day midnight Night IN
'         +
' Next-day morning Night OUT
'=========================================================
Private Sub BuildSessionKeys()

    Dim EmployeeID As Variant
    Dim PunchCollection As Collection
    Dim PunchItem As Variant

    Dim PunchDate As Date
    Dim SessionDate As Date

    Dim DictKey As String

    Set gSessionKeys = CreateObject("Scripting.Dictionary")

    For Each EmployeeID In gEmployeePunches.Keys

        Set PunchCollection = gEmployeePunches(EmployeeID)

        For Each PunchItem In PunchCollection

            PunchDate = DateValue(CDate(PunchItem))

            'Possible attendance date: same date
            SessionDate = PunchDate

            DictKey = CStr(EmployeeID) & "|" & _
                      Format$(SessionDate, "dd-mm-yyyy")

            If Not gSessionKeys.Exists(DictKey) Then
                gSessionKeys.Add DictKey, True
            End If

            'Possible attendance date: previous date
            SessionDate = PunchDate - 1

            DictKey = CStr(EmployeeID) & "|" & _
                      Format$(SessionDate, "dd-mm-yyyy")

            If Not gSessionKeys.Exists(DictKey) Then
                gSessionKeys.Add DictKey, True
            End If

        Next PunchItem

    Next EmployeeID

End Sub

'=========================================================
' PROCESS ALL ATTENDANCE SESSIONS
'=========================================================
Private Sub ProcessAllAttendanceSessions()

    Dim DictKey As Variant

    Dim KeyParts() As String
    Dim DateParts() As String

    Dim EmployeeID As String
    Dim EmployeeName As String
    Dim AttendanceDate As Date

    Dim MorningIn As Date
    Dim MorningOut As Date
    Dim NightIn As Date
    Dim NightOut As Date

    Dim FoundMorningIn As Boolean
    Dim FoundMorningOut As Boolean
    Dim FoundNightIn As Boolean
    Dim FoundNightOut As Boolean

    Dim MorningCount As Long
    Dim NightCount As Long

    For Each DictKey In gSessionKeys.Keys

        KeyParts = Split(CStr(DictKey), "|")

        EmployeeID = KeyParts(0)

        DateParts = Split(KeyParts(1), "-")

        AttendanceDate = DateSerial( _
            CLng(DateParts(2)), _
            CLng(DateParts(1)), _
            CLng(DateParts(0)))

        EmployeeName = vbNullString

        If gEmployeeNames.Exists(EmployeeID) Then
            EmployeeName = gEmployeeNames(EmployeeID)
        End If

        FindSessionPunches _
            EmployeeID, _
            AttendanceDate, _
            MorningIn, _
            MorningOut, _
            NightIn, _
            NightOut, _
            FoundMorningIn, _
            FoundMorningOut, _
            FoundNightIn, _
            FoundNightOut

        '-------------------------------------------------
        ' SHIFT CLASSIFICATION
        '
        ' Morning Shift receives priority when a valid
        ' afternoon punch IN exists.
        '
        ' This prevents the next-day midnight punch from
        ' being incorrectly classified as Night Shift IN.
        '-------------------------------------------------

        If FoundMorningIn Then

            WriteMorningAttendance _
                EmployeeID, _
                EmployeeName, _
                AttendanceDate, _
                MorningIn, _
                MorningOut, _
                FoundMorningOut

            MorningCount = MorningCount + 1

        ElseIf FoundNightIn Or FoundNightOut Then

            WriteNightAttendance _
                EmployeeID, _
                EmployeeName, _
                AttendanceDate, _
                NightIn, _
                NightOut, _
                FoundNightIn, _
                FoundNightOut

            NightCount = NightCount + 1

        End If

    Next DictKey

    SortAttendanceSheets
    FormatAttendanceSheets

        MsgBox _
            "Attendance processing completed successfully." & vbCrLf & _
            "The Raw Punch Database and attendance records have been updated.", _
            vbInformation, _
            "Attendance Automation"

End Sub

'=========================================================
' FIND ALL PUNCH CANDIDATES FOR ONE SESSION
'=========================================================
Private Sub FindSessionPunches( _
    ByVal EmployeeID As String, _
    ByVal AttendanceDate As Date, _
    ByRef MorningIn As Date, _
    ByRef MorningOut As Date, _
    ByRef NightIn As Date, _
    ByRef NightOut As Date, _
    ByRef FoundMorningIn As Boolean, _
    ByRef FoundMorningOut As Boolean, _
    ByRef FoundNightIn As Boolean, _
    ByRef FoundNightOut As Boolean)

    Dim PunchCollection As Collection
    Dim PunchItem As Variant
    Dim FullDateTime As Date

    Dim MorningInStart As Date
    Dim MorningInEnd As Date

    Dim MorningOutStart As Date
    Dim MorningOutEnd As Date

    Dim NightInStart As Date
    Dim NightInEnd As Date

    Dim NightOutStart As Date
    Dim NightOutEnd As Date

    FoundMorningIn = False
    FoundMorningOut = False
    FoundNightIn = False
    FoundNightOut = False

    If Not gEmployeePunches.Exists(EmployeeID) Then Exit Sub

    Set PunchCollection = gEmployeePunches(EmployeeID)

    'Morning Shift IN:
    'Attendance date between 2:00 PM and 4:00 PM
    MorningInStart = AttendanceDate + TimeSerial(14, 0, 0)
    MorningInEnd = AttendanceDate + TimeSerial(16, 0, 0)

    'Morning Shift OUT:
    'Attendance date 6:00 PM until next date 3:00 AM
    MorningOutStart = AttendanceDate + TimeSerial(18, 0, 0)
    MorningOutEnd = AttendanceDate + 1 + TimeSerial(3, 0, 0)

    'Night Shift IN:
    'Next date between 12:00 AM and 3:00 AM
    NightInStart = AttendanceDate + 1
    NightInEnd = AttendanceDate + 1 + TimeSerial(3, 0, 0)

    'Night Shift OUT:
    'Next date between 6:00 AM and 12:00 PM
    NightOutStart = AttendanceDate + 1 + TimeSerial(6, 0, 0)
    NightOutEnd = AttendanceDate + 1 + TimeSerial(12, 0, 0)

    For Each PunchItem In PunchCollection

        FullDateTime = CDate(PunchItem)

        'Earliest Morning IN
        If FullDateTime >= MorningInStart And _
           FullDateTime <= MorningInEnd Then

            If Not FoundMorningIn Then

                MorningIn = FullDateTime
                FoundMorningIn = True

            ElseIf FullDateTime < MorningIn Then

                MorningIn = FullDateTime

            End If

        End If

        'Latest Morning OUT
        If FullDateTime >= MorningOutStart And _
           FullDateTime <= MorningOutEnd Then

            If Not FoundMorningOut Then

                MorningOut = FullDateTime
                FoundMorningOut = True

            ElseIf FullDateTime > MorningOut Then

                MorningOut = FullDateTime

            End If

        End If

        'Earliest Night IN
        If FullDateTime >= NightInStart And _
           FullDateTime <= NightInEnd Then

            If Not FoundNightIn Then

                NightIn = FullDateTime
                FoundNightIn = True

            ElseIf FullDateTime < NightIn Then

                NightIn = FullDateTime

            End If

        End If

        'Latest Night OUT
        If FullDateTime >= NightOutStart And _
           FullDateTime <= NightOutEnd Then

            If Not FoundNightOut Then

                NightOut = FullDateTime
                FoundNightOut = True

            ElseIf FullDateTime > NightOut Then

                NightOut = FullDateTime

            End If

        End If

    Next PunchItem

End Sub

'=========================================================
' WRITE MORNING SHIFT ATTENDANCE
'=========================================================
Private Sub WriteMorningAttendance( _
    ByVal EmployeeID As String, _
    ByVal EmployeeName As String, _
    ByVal AttendanceDate As Date, _
    ByVal PunchIn As Date, _
    ByVal PunchOut As Date, _
    ByVal FoundPunchOut As Boolean)

    Dim ws As Worksheet
    Dim NextRow As Long

    Dim TotalHours As Double
    Dim Status As String

    Set ws = ThisWorkbook.Worksheets("Morning Shift Database")

    NextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    If NextRow < 2 Then NextRow = 2

    If FoundPunchOut Then

        TotalHours = CalculateHours(PunchIn, PunchOut)
        Status = "Present"

    Else

        TotalHours = 0
        Status = "UNK"

    End If

    ws.Cells(NextRow, 1).Value = AttendanceDate
    ws.Cells(NextRow, 2).Value = EmployeeID
    ws.Cells(NextRow, 3).Value = EmployeeName
    ws.Cells(NextRow, 4).Value = PunchIn

    If FoundPunchOut Then
        ws.Cells(NextRow, 5).Value = PunchOut
    End If

    ws.Cells(NextRow, 6).Value = TotalHours
    ws.Cells(NextRow, 7).Value = "Morning"
    ws.Cells(NextRow, 8).Value = Status

End Sub

'=========================================================
' WRITE NIGHT SHIFT ATTENDANCE
'=========================================================
Private Sub WriteNightAttendance( _
    ByVal EmployeeID As String, _
    ByVal EmployeeName As String, _
    ByVal AttendanceDate As Date, _
    ByVal PunchIn As Date, _
    ByVal PunchOut As Date, _
    ByVal FoundPunchIn As Boolean, _
    ByVal FoundPunchOut As Boolean)

    Dim ws As Worksheet
    Dim NextRow As Long

    Dim TotalHours As Double
    Dim Status As String

    Set ws = ThisWorkbook.Worksheets("Night Shift Database")

    NextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    If NextRow < 2 Then NextRow = 2

    If FoundPunchIn And FoundPunchOut Then

        TotalHours = CalculateHours(PunchIn, PunchOut)
        Status = "Present"

    Else

        TotalHours = 0
        Status = "UNK"

    End If

    ws.Cells(NextRow, 1).Value = AttendanceDate
    ws.Cells(NextRow, 2).Value = EmployeeID
    ws.Cells(NextRow, 3).Value = EmployeeName

    If FoundPunchIn Then
        ws.Cells(NextRow, 4).Value = PunchIn
    End If

    If FoundPunchOut Then
        ws.Cells(NextRow, 5).Value = PunchOut
    End If

    ws.Cells(NextRow, 6).Value = TotalHours
    ws.Cells(NextRow, 7).Value = "Night"
    ws.Cells(NextRow, 8).Value = Status

End Sub

'=========================================================
' FORMAT ATTENDANCE DATABASES
'=========================================================
Private Sub FormatAttendanceSheets()

    Dim wsMorning As Worksheet
    Dim wsNight As Worksheet

    Set wsMorning = _
        ThisWorkbook.Worksheets("Morning Shift Database")

    Set wsNight = _
        ThisWorkbook.Worksheets("Night Shift Database")

    FormatOneAttendanceSheet wsMorning
    FormatOneAttendanceSheet wsNight

End Sub

Private Sub FormatOneAttendanceSheet(ByVal ws As Worksheet)

    Dim LastRow As Long

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If LastRow < 2 Then Exit Sub

    ws.Range("A2:A" & LastRow).NumberFormat = "dd-mmm-yyyy"

    ws.Range("D2:E" & LastRow).NumberFormat = _
        "dd-mmm-yyyy hh:mm:ss AM/PM"

    ws.Range("F2:F" & LastRow).NumberFormat = "0.00"

    ws.Columns("A:H").AutoFit

End Sub

'=========================================================
' SORT BOTH ATTENDANCE DATABASES
'=========================================================
Private Sub SortAttendanceSheets()

    SortOneAttendanceSheet _
        ThisWorkbook.Worksheets("Morning Shift Database")

    SortOneAttendanceSheet _
        ThisWorkbook.Worksheets("Night Shift Database")

End Sub

'=========================================================
' SORT ONE ATTENDANCE DATABASE
'
' Sort order:
' 1. Attendance Date
' 2. Punch IN
' 3. Employee ID
'=========================================================
Private Sub SortOneAttendanceSheet(ByVal ws As Worksheet)

    Dim LastRow As Long

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If LastRow < 3 Then Exit Sub

    With ws.Sort

        .SortFields.Clear

        .SortFields.Add _
            Key:=ws.Range("A2:A" & LastRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SortFields.Add _
            Key:=ws.Range("D2:D" & LastRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SortFields.Add _
            Key:=ws.Range("B2:B" & LastRow), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        .SetRange ws.Range("A1:H" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply

    End With

End Sub

