Attribute VB_Name = "modSummary"
Option Explicit

Public Sub UpdateTodaysSummary()

    Dim wsSummary As Worksheet
    Dim wsRawImport As Worksheet
    Dim wsMorning As Worksheet
    Dim wsNight As Worksheet

    Dim ReportingDate As Date
    Dim PreviousNightDate As Date

    Dim MorningCount As Long
    Dim NightCount As Long
    Dim MorningIncomplete As Long
    Dim NightIncomplete As Long
    Dim UnknownCount As Long
    Dim TotalReportedManpower As Long

    On Error GoTo ErrorHandler

    Set wsSummary = ThisWorkbook.Worksheets("Todays Summary")
    Set wsRawImport = ThisWorkbook.Worksheets("Raw Import")
    Set wsMorning = ThisWorkbook.Worksheets("Morning Shift Database")
    Set wsNight = ThisWorkbook.Worksheets("Night Shift Database")

    ReportingDate = GetLatestRawImportDate(wsRawImport)

    If ReportingDate = 0 Then
        MsgBox _
            "The reporting date could not be identified from Raw Import.", _
            vbExclamation, _
            "Summary Update"

        Exit Sub
    End If

    PreviousNightDate = ReportingDate - 1

    MorningCount = CountMorningShiftEmployees( _
                        wsMorning, _
                        ReportingDate)

    MorningIncomplete = CountMorningIncompleteRecords( _
                            wsMorning, _
                            ReportingDate)

    NightCount = CountCompletedNightShiftEmployees( _
                     wsNight, _
                     PreviousNightDate)

    NightIncomplete = CountNightIncompleteRecords( _
                          wsNight, _
                          PreviousNightDate)

    UnknownCount = CountUnknownRecords( _
                       wsMorning, _
                       wsNight, _
                       ReportingDate, _
                       PreviousNightDate)

TotalReportedManpower = MorningCount + NightCount

'Build and format the summary layout first.
FormatSummarySheet wsSummary

'Write the calculated values after the layout is created.
WriteSummaryValues _
    wsSummary, _
    ReportingDate, _
    MorningCount, _
    NightCount, _
    TotalReportedManpower, _
    MorningIncomplete, _
    NightIncomplete, _
    UnknownCount

Exit Sub

ErrorHandler:

    MsgBox _
        "The management summary could not be updated." & vbCrLf & _
        "Error: " & Err.Description, _
        vbCritical, _
        "Summary Error"

End Sub


Private Function GetLatestRawImportDate( _
    ByVal ws As Worksheet) As Date

    Dim LastRow As Long
    Dim RowNum As Long
    Dim DateColumn As Long
    Dim CurrentDate As Variant
    Dim LatestDate As Date

    DateColumn = gHeaderMap("Date")

    LastRow = ws.Cells(ws.Rows.Count, DateColumn).End(xlUp).Row

    For RowNum = DATA_START_ROW To LastRow

        CurrentDate = ws.Cells(RowNum, DateColumn).Value

        If IsDate(CurrentDate) Then

            If DateValue(CurrentDate) > LatestDate Then
                LatestDate = DateValue(CurrentDate)
            End If

        End If

    Next RowNum

    GetLatestRawImportDate = LatestDate

End Function


Private Function CountMorningShiftEmployees( _
    ByVal ws As Worksheet, _
    ByVal AttendanceDate As Date) As Long

    Dim LastRow As Long
    Dim RowNum As Long
    Dim RecordDate As Variant
    Dim PunchInValue As Variant

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For RowNum = DATA_START_ROW To LastRow

        RecordDate = ws.Cells(RowNum, 1).Value
        PunchInValue = ws.Cells(RowNum, 4).Value

        If IsDate(RecordDate) Then

            If DateValue(RecordDate) = AttendanceDate Then

                If HasValidValue(PunchInValue) Then
                    CountMorningShiftEmployees = _
                        CountMorningShiftEmployees + 1
                End If

            End If

        End If

    Next RowNum

End Function


Private Function CountMorningIncompleteRecords( _
    ByVal ws As Worksheet, _
    ByVal AttendanceDate As Date) As Long

    Dim LastRow As Long
    Dim RowNum As Long
    Dim RecordDate As Variant
    Dim PunchInValue As Variant
    Dim PunchOutValue As Variant

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For RowNum = DATA_START_ROW To LastRow

        RecordDate = ws.Cells(RowNum, 1).Value
        PunchInValue = ws.Cells(RowNum, 4).Value
        PunchOutValue = ws.Cells(RowNum, 5).Value

        If IsDate(RecordDate) Then

            If DateValue(RecordDate) = AttendanceDate Then

                If HasValidValue(PunchInValue) _
                   And Not HasValidValue(PunchOutValue) Then

                    CountMorningIncompleteRecords = _
                        CountMorningIncompleteRecords + 1

                End If

            End If

        End If

    Next RowNum

End Function


Private Function CountCompletedNightShiftEmployees( _
    ByVal ws As Worksheet, _
    ByVal AttendanceDate As Date) As Long

    Dim LastRow As Long
    Dim RowNum As Long
    Dim RecordDate As Variant
    Dim PunchInValue As Variant
    Dim PunchOutValue As Variant

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For RowNum = DATA_START_ROW To LastRow

        RecordDate = ws.Cells(RowNum, 1).Value
        PunchInValue = ws.Cells(RowNum, 4).Value
        PunchOutValue = ws.Cells(RowNum, 5).Value

        If IsDate(RecordDate) Then

            If DateValue(RecordDate) = AttendanceDate Then

                If HasValidValue(PunchInValue) _
                   And HasValidValue(PunchOutValue) Then

                    CountCompletedNightShiftEmployees = _
                        CountCompletedNightShiftEmployees + 1

                End If

            End If

        End If

    Next RowNum

End Function


Private Function CountNightIncompleteRecords( _
    ByVal ws As Worksheet, _
    ByVal AttendanceDate As Date) As Long

    Dim LastRow As Long
    Dim RowNum As Long
    Dim RecordDate As Variant
    Dim PunchInValue As Variant
    Dim PunchOutValue As Variant

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For RowNum = DATA_START_ROW To LastRow

        RecordDate = ws.Cells(RowNum, 1).Value
        PunchInValue = ws.Cells(RowNum, 4).Value
        PunchOutValue = ws.Cells(RowNum, 5).Value

        If IsDate(RecordDate) Then

            If DateValue(RecordDate) = AttendanceDate Then

                If HasValidValue(PunchInValue) _
                   Xor HasValidValue(PunchOutValue) Then

                    CountNightIncompleteRecords = _
                        CountNightIncompleteRecords + 1

                End If

            End If

        End If

    Next RowNum

End Function


Private Function CountUnknownRecords( _
    ByVal wsMorning As Worksheet, _
    ByVal wsNight As Worksheet, _
    ByVal MorningDate As Date, _
    ByVal NightDate As Date) As Long

    CountUnknownRecords = _
        CountStatusRecords(wsMorning, MorningDate, "UNK") + _
        CountStatusRecords(wsNight, NightDate, "UNK")

End Function


Private Function CountStatusRecords( _
    ByVal ws As Worksheet, _
    ByVal AttendanceDate As Date, _
    ByVal RequiredStatus As String) As Long

    Dim LastRow As Long
    Dim RowNum As Long
    Dim RecordDate As Variant
    Dim StatusValue As String

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For RowNum = DATA_START_ROW To LastRow

        RecordDate = ws.Cells(RowNum, 1).Value
        StatusValue = Trim$(CStr(ws.Cells(RowNum, 8).Value))

        If IsDate(RecordDate) Then

            If DateValue(RecordDate) = AttendanceDate _
               And StrComp( _
                       StatusValue, _
                       RequiredStatus, _
                       vbTextCompare) = 0 Then

                CountStatusRecords = CountStatusRecords + 1

            End If

        End If

    Next RowNum

End Function


Private Function HasValidValue( _
    ByVal CellValue As Variant) As Boolean

    If IsError(CellValue) Then
        Exit Function
    End If

    If Len(Trim$(CStr(CellValue))) = 0 Then
        Exit Function
    End If

    HasValidValue = True

End Function


Private Sub WriteSummaryValues( _
    ByVal ws As Worksheet, _
    ByVal ReportingDate As Date, _
    ByVal MorningCount As Long, _
    ByVal NightCount As Long, _
    ByVal TotalManpower As Long, _
    ByVal MorningIncomplete As Long, _
    ByVal NightIncomplete As Long, _
    ByVal UnknownCount As Long)

    'Attendance summary
    ws.Range("B5").Value = ReportingDate
    ws.Range("B6").Value = MorningCount
    ws.Range("B7").Value = NightCount
    ws.Range("B8").Value = TotalManpower

    ws.Range("B11").Value = MorningIncomplete
    ws.Range("B12").Value = NightIncomplete
    ws.Range("B13").Value = UnknownCount

    'System information
    ws.Range("B16").Value = GetSelectedACSFileName()
    ws.Range("B17").Value = Date
    ws.Range("B18").Value = Time
    ws.Range("B19").Value = "Completed Successfully"

End Sub


Private Function GetSelectedACSFileName() As String

    Dim FileNamePosition As Long

    If Len(Trim$(gACSFilePath)) = 0 Then
        GetSelectedACSFileName = "Not Available"
        Exit Function
    End If

    FileNamePosition = InStrRev(gACSFilePath, Application.PathSeparator)

    If FileNamePosition > 0 Then

        GetSelectedACSFileName = _
            Mid$(gACSFilePath, FileNamePosition + 1)

    Else

        GetSelectedACSFileName = gACSFilePath

    End If

End Function


Private Sub FormatSummarySheet(ByVal ws As Worksheet)

    With ws

        Application.DisplayAlerts = False

        'Remove previous merges and formatting
        .Cells.UnMerge
        .Range("A1:B20").ClearFormats
        .Range("A1:B20").ClearContents

        Application.DisplayAlerts = True

        'Main titles
        .Range("A1:B1").Merge
        .Range("A2:B2").Merge

        .Range("A1").Value = _
            "POLATI ATTENDANCE AUTOMATION SYSTEM"

        .Range("A2").Value = _
            "DAILY MANAGEMENT SUMMARY"

        'Attendance section
        .Range("A4:B4").Merge
        .Range("A4").Value = "ATTENDANCE SUMMARY"

        .Range("A5").Value = "Reporting Date"
        .Range("A6").Value = "Today's Morning Shift"
        .Range("A7").Value = "Previous Night Shift"
        .Range("A8").Value = "Total Reported Manpower"

        'Attendance quality section
        .Range("A10:B10").Merge
        .Range("A10").Value = "ATTENDANCE QUALITY"

        .Range("A11").Value = "Morning Incomplete"
        .Range("A12").Value = "Night Incomplete"
        .Range("A13").Value = "Unknown Records"

        'System information section
        .Range("A15:B15").Merge
        .Range("A15").Value = "SYSTEM INFORMATION"

        .Range("A16").Value = "Last Imported ACS File"
        .Range("A17").Value = "Last Processed Date"
        .Range("A18").Value = "Last Processed Time"
        .Range("A19").Value = "Processing Status"

        'Column sizes
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 42

        'Row heights
        .Rows(1).RowHeight = 32
        .Rows(2).RowHeight = 25
        .Rows("4:19").RowHeight = 23

        'Main title formatting
        With .Range("A1:B1")

            .Interior.Color = RGB(16, 42, 67)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 18
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

        End With

        'Subtitle formatting
        With .Range("A2:B2")

            .Interior.Color = RGB(217, 217, 217)
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter

        End With

        'Section header formatting
        FormatSummarySectionHeader ws.Range("A4:B4")
        FormatSummarySectionHeader ws.Range("A10:B10")
        FormatSummarySectionHeader ws.Range("A15:B15")

        'Label formatting
        With .Range("A5:A19")

            .Font.Bold = True
            .Font.Size = 10
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter

        End With

        'General borders
        With .Range("A5:B8")

            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
            .Borders.Weight = xlThin

        End With

        With .Range("A11:B13")

            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
            .Borders.Weight = xlThin

        End With

        With .Range("A16:B19")

            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
            .Borders.Weight = xlThin

        End With

        'Value alignment
        .Range("B5:B19").HorizontalAlignment = xlCenter
        .Range("B5:B19").VerticalAlignment = xlCenter

        'Date and time formats
        .Range("B5").NumberFormat = "dd-mmm-yyyy"
        .Range("B17").NumberFormat = "dd-mmm-yyyy"
        .Range("B18").NumberFormat = "hh:mm:ss AM/PM"

        'Main KPI values
        With .Range("B6:B7")

            .Font.Bold = True
            .Font.Size = 17
            .Font.Color = RGB(0, 51, 102)

        End With

        With .Range("B8")

            .Font.Bold = True
            .Font.Size = 19
            .Font.Color = RGB(16, 42, 67)

        End With

        'Incomplete and unknown values
        With .Range("B11:B13")

            .Font.Bold = True
            .Font.Size = 13
            .Font.Color = RGB(192, 80, 0)

        End With

        'System information
        With .Range("B16:B19")

            .Font.Size = 10
            .HorizontalAlignment = xlLeft

        End With

        'Processing status
        With .Range("B19")

            .Font.Bold = True
            .Font.Color = RGB(0, 112, 0)
            .HorizontalAlignment = xlCenter

        End With

        'Backgrounds
        .Range("A5:B8").Interior.Color = RGB(248, 248, 248)
        .Range("A11:B13").Interior.Color = RGB(248, 248, 248)
        .Range("A16:B19").Interior.Color = RGB(248, 248, 248)

    End With

End Sub

Private Sub FormatSummarySectionHeader(ByVal HeaderRange As Range)

    With HeaderRange

        .Interior.Color = RGB(16, 42, 67)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

End Sub

