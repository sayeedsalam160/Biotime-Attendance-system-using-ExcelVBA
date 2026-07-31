Attribute VB_Name = "modNightShift"
Option Explicit

Public Sub ProcessNightShift()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim RowNum As Long

    Dim EmployeeID As String
    Dim PunchDate As Date
    Dim PunchTime As Date
    Dim AttendanceDate As Date
    Dim FullDateTime As Date
    Dim DictKey As String

    Dim RawDate As Variant
    Dim RawTime As Variant

    Dim PunchCollection As Collection

    Dim NightInCount As Long
    Dim NightOutCount As Long

    Set ws = ThisWorkbook.Worksheets("Raw Punch Database")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set gNightShiftDict = CreateObject("Scripting.Dictionary")

    For RowNum = DATA_START_ROW To LastRow

        EmployeeID = Trim$(CStr( _
            ws.Cells(RowNum, gHeaderMap("Employee ID")).Value))

        RawDate = ws.Cells( _
            RowNum, gHeaderMap("Date")).Value

        RawTime = ws.Cells( _
            RowNum, gHeaderMap("Time")).Value

        If Len(EmployeeID) > 0 Then

            'The Date column must contain a valid Excel date.
            If IsDate(RawDate) Or IsNumeric(RawDate) Then

                'The Time column contains an Excel numeric time value.
                If IsDate(RawTime) Or IsNumeric(RawTime) Then

                    PunchDate = DateValue(CDate(RawDate))

                    'Extract only the time portion.
                    PunchTime = TimeValue(CDate(RawTime))

                    If IsNightIn(PunchTime) Then
                        NightInCount = NightInCount + 1
                    End If

                    If IsNightOut(PunchTime) Then
                        NightOutCount = NightOutCount + 1
                    End If

                    If IsNightIn(PunchTime) Or _
                       IsNightOut(PunchTime) Then

                        FullDateTime = PunchDate + PunchTime

                        'Both night IN and night OUT belong
                        'to the previous attendance date.
                        AttendanceDate = PunchDate - 1

                        DictKey = EmployeeID & "|" & _
                                  Format$( _
                                      AttendanceDate, _
                                      "dd-mm-yyyy")

                        If Not gNightShiftDict.Exists(DictKey) Then

                            Set PunchCollection = New Collection

                            PunchCollection.Add FullDateTime

                            gNightShiftDict.Add _
                                DictKey, PunchCollection

                        Else

                            gNightShiftDict(DictKey).Add _
                                FullDateTime

                        End If

                    End If

                End If

            End If

        End If

    Next RowNum

    Debug.Print "Night IN Punches = " & NightInCount
    Debug.Print "Night OUT Punches = " & NightOutCount
    Debug.Print "Night Dictionary Count = " & _
                gNightShiftDict.Count

End Sub

Public Sub WriteNightShiftDatabase()

    Dim wsDB As Worksheet
    Dim DBRow As Long
    Dim WrittenCount As Long

    Dim DictKey As Variant
    Dim PunchCollection As Collection
    Dim PunchItem As Variant

    Dim EmployeeID As String
    Dim AttendanceDate As Date
    Dim KeyParts() As String
    Dim DateParts() As String

    Dim PunchIn As Date
    Dim PunchOut As Date

    Dim FoundIn As Boolean
    Dim FoundOut As Boolean

    Dim TotalHours As Double
    Dim Status As String

    Set wsDB = ThisWorkbook.Worksheets("Night Shift Database")

    If gNightShiftDict Is Nothing Then

        MsgBox "Night Shift dictionary is not initialized.", _
               vbExclamation

        Exit Sub

    End If

    DBRow = 2
    WrittenCount = 0

    For Each DictKey In gNightShiftDict.Keys

        Set PunchCollection = gNightShiftDict(DictKey)

        FoundIn = False
        FoundOut = False
        TotalHours = 0

        KeyParts = Split(CStr(DictKey), "|")

        EmployeeID = KeyParts(0)

        DateParts = Split(KeyParts(1), "-")

        AttendanceDate = DateSerial( _
            CLng(DateParts(2)), _
            CLng(DateParts(1)), _
            CLng(DateParts(0)))

        For Each PunchItem In PunchCollection

            If IsNightIn(CDate(PunchItem)) Then

                If Not FoundIn Then

                    PunchIn = CDate(PunchItem)
                    FoundIn = True

                ElseIf CDate(PunchItem) < PunchIn Then

                    PunchIn = CDate(PunchItem)

                End If

            End If

            If IsNightOut(CDate(PunchItem)) Then

                If Not FoundOut Then

                    PunchOut = CDate(PunchItem)
                    FoundOut = True

                ElseIf CDate(PunchItem) > PunchOut Then

                    PunchOut = CDate(PunchItem)

                End If

            End If

        Next PunchItem

            'Case 1: Both Night IN and Night OUT exist.
            If FoundIn And FoundOut Then
            
                TotalHours = CalculateHours(PunchIn, PunchOut)
                Status = "Present"
            
                wsDB.Cells(DBRow, 1).Value = AttendanceDate
                wsDB.Cells(DBRow, 2).Value = EmployeeID
                wsDB.Cells(DBRow, 4).Value = PunchIn
                wsDB.Cells(DBRow, 5).Value = PunchOut
                wsDB.Cells(DBRow, 6).Value = TotalHours
                wsDB.Cells(DBRow, 7).Value = "Night"
                wsDB.Cells(DBRow, 8).Value = Status
            
                DBRow = DBRow + 1
                WrittenCount = WrittenCount + 1
            
            'Case 2: Night OUT exists, but Night IN is missing.
            ElseIf FoundOut Then
            
                Status = "UNK"
            
                wsDB.Cells(DBRow, 1).Value = AttendanceDate
                wsDB.Cells(DBRow, 2).Value = EmployeeID
                wsDB.Cells(DBRow, 5).Value = PunchOut
                wsDB.Cells(DBRow, 6).Value = 0
                wsDB.Cells(DBRow, 7).Value = "Night"
                wsDB.Cells(DBRow, 8).Value = Status
            
                DBRow = DBRow + 1
                WrittenCount = WrittenCount + 1
            
            'Case 3: Midnight punch only.
            ElseIf FoundIn Then
            
                'Exclude it from Night Shift when the employee has a valid
                'Morning Shift punch-in for the same attendance date.
                If Not HasMorningShiftPunchIn(EmployeeID, AttendanceDate) Then
            
                    Status = "UNK"
            
                    wsDB.Cells(DBRow, 1).Value = AttendanceDate
                    wsDB.Cells(DBRow, 2).Value = EmployeeID
                    wsDB.Cells(DBRow, 4).Value = PunchIn
                    wsDB.Cells(DBRow, 6).Value = 0
                    wsDB.Cells(DBRow, 7).Value = "Night"
                    wsDB.Cells(DBRow, 8).Value = Status
            
                    DBRow = DBRow + 1
                    WrittenCount = WrittenCount + 1
            
                End If
            
            End If

    Next DictKey

    FormatNightShiftDatabase wsDB

    MsgBox WrittenCount & _
           " Night Shift records written successfully.", _
           vbInformation

End Sub


Public Function IsNightIn(ByVal PunchTime As Date) As Boolean

    Dim T As Date

    T = TimeValue(PunchTime)

    IsNightIn = (T >= TimeSerial(0, 0, 0) And _
                 T <= TimeSerial(3, 0, 0))

End Function


Public Function IsNightOut(ByVal PunchTime As Date) As Boolean

    Dim T As Date

    T = TimeValue(PunchTime)

    IsNightOut = (T >= TimeSerial(6, 0, 0) And _
                  T <= TimeSerial(12, 0, 0))

End Function


Private Sub FormatNightShiftDatabase(ByVal ws As Worksheet)

    ws.Columns(1).NumberFormat = "dd-mmm-yyyy"
    ws.Columns(4).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
    ws.Columns(5).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
    ws.Columns(6).NumberFormat = "0.00"
    ws.Columns("A:H").AutoFit

End Sub

Private Function HasMorningShiftPunchIn( _
    ByVal EmployeeID As String, _
    ByVal AttendanceDate As Date) As Boolean

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim RowNum As Long

    Set ws = ThisWorkbook.Worksheets("Morning Shift Database")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For RowNum = 2 To LastRow

        If Trim$(CStr(ws.Cells(RowNum, 2).Value)) = EmployeeID Then

            If IsDate(ws.Cells(RowNum, 1).Value) Then

                If DateValue(ws.Cells(RowNum, 1).Value) = _
                   DateValue(AttendanceDate) Then

                    If Len(ws.Cells(RowNum, 4).Value) > 0 Then
                        HasMorningShiftPunchIn = True
                        Exit Function
                    End If

                End If

            End If

        End If

    Next RowNum

    HasMorningShiftPunchIn = False

End Function

