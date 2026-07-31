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
    Dim FullDateTime As Date

    Dim PunchCollection As Collection

    Set ws = ThisWorkbook.Worksheets("Raw Punch Database")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set gMorningShiftDict = CreateObject("Scripting.Dictionary")

    For RowNum = 2 To LastRow

        EmployeeID = Trim(ws.Cells(RowNum, gHeaderMap("Employee ID")).Value)

        PunchDate = ws.Cells(RowNum, gHeaderMap("Date")).Value

        PunchTime = ws.Cells(RowNum, gHeaderMap("Time")).Value

        FullDateTime = PunchDate + TimeValue(PunchTime)

        AttendanceDate = GetAttendanceDate(PunchDate, PunchTime)

        DictKey = EmployeeID & "|" & Format(AttendanceDate, "dd-mm-yyyy")

        If Not gMorningShiftDict.Exists(DictKey) Then

            Set PunchCollection = New Collection

            PunchCollection.Add FullDateTime

            gMorningShiftDict.Add DictKey, PunchCollection

        Else

            gMorningShiftDict(DictKey).Add FullDateTime

        End If

    Next RowNum

    Debug.Print "Morning Dictionary Count = " & gMorningShiftDict.Count

End Sub

Public Sub WriteMorningShiftDatabase()

    Dim wsDB As Worksheet
    Dim DBRow As Long

    Dim DictKey As Variant
    Dim PunchCollection As Collection
    Dim PunchItem As Variant

    Dim EmployeeID As String
    Dim AttendanceDate As Date

    Dim PunchIn As Variant
    Dim PunchOut As Variant

    Dim FoundIn As Boolean
    Dim FoundOut As Boolean

    Dim TotalHours As Double
    Dim Status As String

    Dim Parts() As String

    Set wsDB = ThisWorkbook.Worksheets("Morning Shift Database")

    For Each DictKey In gMorningShiftDict.Keys

        Set PunchCollection = gMorningShiftDict(DictKey)

        FoundIn = False
        FoundOut = False

        Parts = Split(DictKey, "|")

        EmployeeID = Parts(0)
        AttendanceDate = CDate(Parts(1))

        For Each PunchItem In PunchCollection

            If IsMorningIn(PunchItem) Then

                If Not FoundIn Then

                    PunchIn = PunchItem
                    FoundIn = True

                ElseIf PunchItem < PunchIn Then

                    PunchIn = PunchItem

                End If

            End If

            If IsMorningOut(PunchItem) Then

                If Not FoundOut Then

                    PunchOut = PunchItem
                    FoundOut = True

                ElseIf PunchItem > PunchOut Then

                    PunchOut = PunchItem

                End If

            End If

        Next PunchItem

        If FoundIn And FoundOut Then

            TotalHours = CalculateHours(PunchIn, PunchOut)

            Status = "Present"

        Else

            TotalHours = 0

            Status = "UNK"

        End If

        If Not RecordExists(EmployeeID, AttendanceDate, _
                            "Morning Shift Database") Then

            DBRow = GetNextRow("Morning Shift Database")

            wsDB.Cells(DBRow, 1).Value = AttendanceDate
            wsDB.Cells(DBRow, 2).Value = EmployeeID

            If FoundIn Then
                wsDB.Cells(DBRow, 4).Value = PunchIn
            End If

            If FoundOut Then
                wsDB.Cells(DBRow, 5).Value = PunchOut
            End If

            wsDB.Cells(DBRow, 6).Value = TotalHours
            wsDB.Cells(DBRow, 7).Value = "Morning"
            wsDB.Cells(DBRow, 8).Value = Status

        End If

    Next DictKey

    MsgBox "Morning Shift Database Updated", vbInformation

End Sub

Public Function IsMorningIn(ByVal PunchTime As Date) As Boolean

    Dim T As Date

    T = TimeValue(PunchTime)

    IsMorningIn = (T >= TimeValue("14:00:00") And _
                   T <= TimeValue("16:00:00"))

End Function

Public Function IsMorningOut(ByVal PunchTime As Date) As Boolean

    Dim T As Date

    T = TimeValue(PunchTime)

    IsMorningOut = _
        (T >= TimeValue("18:00:00") And _
         T <= TimeValue("23:59:59")) _
         Or _
        (T >= TimeValue("00:00:00") And _
         T <= TimeValue("02:59:59"))

End Function

