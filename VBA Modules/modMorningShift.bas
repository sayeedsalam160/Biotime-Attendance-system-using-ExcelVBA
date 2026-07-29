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
    Dim FullDateTime As Date
    Dim TestKey As Variant

    Set ws = ThisWorkbook.Worksheets("Raw Import")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Set gMorningShiftDict = CreateObject("Scripting.Dictionary")

    For RowNum = DATA_START_ROW To LastRow

        EmployeeID = Trim(ws.Cells(RowNum, gHeaderMap("Employee ID")).Value)

        PunchDate = ws.Cells(RowNum, gHeaderMap("Date")).Value

        PunchTime = ws.Cells(RowNum, gHeaderMap("Time")).Value
        
        FullDateTime = PunchDate + TimeValue(PunchTime)

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
        
            gMorningShiftDict.Add DictKey, Array(FullDateTime, FullDateTime)
        
        Else
        
            PunchData = gMorningShiftDict(DictKey)
        
            If FullDateTime < PunchData(0) Then
                PunchData(0) = FullDateTime
            End If
        
            If FullDateTime > PunchData(1) Then
                PunchData(1) = FullDateTime
            End If
        
            gMorningShiftDict(DictKey) = PunchData
        
        End If

    Next RowNum

    Debug.Print "Last Row = " & LastRow
    Debug.Print "Dictionary Count = " & gMorningShiftDict.Count
    

For Each TestKey In gMorningShiftDict.Keys

    If InStr(TestKey, "30008") > 0 Then

        PunchData = gMorningShiftDict(TestKey)

        Debug.Print "FOUND: " & TestKey
        Debug.Print "First = " & PunchData(0)
        Debug.Print "Last  = " & PunchData(1)
        Debug.Print "--------------------"

    End If

Next TestKey
    Debug.Print "=============================="

    Counter = 0

    For Each Key In gMorningShiftDict.Keys

        PunchData = gMorningShiftDict(Key)

        Debug.Print Key & _
            " | First=" & Format(PunchData(0), "dd-mm-yyyy hh:mm:ss AM/PM") & _
            " | Last=" & Format(PunchData(1), "dd-mm-yyyy hh:mm:ss AM/PM")

        Counter = Counter + 1

        If Counter >= 10 Then Exit For

    Next Key

    MsgBox gMorningShiftDict.Count & _
           " unique Employee-Date combinations found.", vbInformation

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

