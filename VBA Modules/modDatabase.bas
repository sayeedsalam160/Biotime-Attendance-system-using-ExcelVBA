Attribute VB_Name = "modDatabase"
Option Explicit

Public Function RecordExists(ByVal EmployeeID As String, _
                             ByVal AttendanceDate As Date, _
                             ByVal SheetName As String) As Boolean

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim RowNum As Long
    Dim Key As String
    Dim ExistingKey As String

    Set ws = ThisWorkbook.Worksheets(SheetName)

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Key = EmployeeID & "|" & Format(AttendanceDate, "dd-mm-yyyy")

    For RowNum = 2 To LastRow

        ExistingKey = Trim(ws.Cells(RowNum, 2).Value) & "|" & _
                      Format(ws.Cells(RowNum, 1).Value, "dd-mm-yyyy")

        If ExistingKey = Key Then

            RecordExists = True
            Exit Function

        End If

    Next RowNum

    RecordExists = False

End Function


Public Function GetNextRow(ByVal SheetName As String) As Long

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SheetName)

    GetNextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

End Function

Public Function PunchExists(ByVal EmployeeID As String, _
                            ByVal PunchDate As Date, _
                            ByVal PunchTime As Date) As Boolean

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim RowNum As Long

    Set ws = ThisWorkbook.Worksheets("Raw Punch Database")

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For RowNum = 2 To LastRow

        If CStr(ws.Cells(RowNum, 1).Value) = CStr(EmployeeID) _
           And ws.Cells(RowNum, 4).Value = PunchDate _
           And Format(ws.Cells(RowNum, 5).Value, "hh:mm:ss") = _
               Format(PunchTime, "hh:mm:ss") Then

            PunchExists = True
            Exit Function

        End If

    Next RowNum

    PunchExists = False

End Function

Public Sub UpdateRawPunchDatabase()

    Dim wsRaw As Worksheet
    Dim wsDB As Worksheet

    Dim LastRow As Long
    Dim DBRow As Long
    Dim RowNum As Long

    Dim EmployeeID As String
    Dim PunchDate As Date
    Dim PunchTime As Date

    Set wsRaw = ThisWorkbook.Worksheets("Raw Import")
    Set wsDB = ThisWorkbook.Worksheets("Raw Punch Database")

    LastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    For RowNum = 2 To LastRow

        EmployeeID = Trim(wsRaw.Cells(RowNum, 1).Value)

        PunchDate = wsRaw.Cells(RowNum, 4).Value

        PunchTime = wsRaw.Cells(RowNum, 5).Value

        If Not PunchExists(EmployeeID, PunchDate, PunchTime) Then

            DBRow = GetNextRow("Raw Punch Database")

            wsDB.Cells(DBRow, 1).Value = wsRaw.Cells(RowNum, 1).Value
            wsDB.Cells(DBRow, 2).Value = wsRaw.Cells(RowNum, 2).Value
            wsDB.Cells(DBRow, 3).Value = wsRaw.Cells(RowNum, 3).Value
            wsDB.Cells(DBRow, 4).Value = wsRaw.Cells(RowNum, 4).Value
            wsDB.Cells(DBRow, 5).Value = wsRaw.Cells(RowNum, 5).Value
            wsDB.Cells(DBRow, 6).Value = wsRaw.Cells(RowNum, 6).Value
            wsDB.Cells(DBRow, 7).Value = wsRaw.Cells(RowNum, 7).Value
            wsDB.Cells(DBRow, 8).Value = wsRaw.Cells(RowNum, 8).Value
            wsDB.Cells(DBRow, 9).Value = wsRaw.Cells(RowNum, 9).Value
            wsDB.Cells(DBRow, 10).Value = wsRaw.Cells(RowNum, 10).Value

        End If

    Next RowNum

    MsgBox "Raw Punch Database updated successfully.", vbInformation

End Sub



Public Sub ClearAttendanceDatabase(ByVal SheetName As String)

    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Worksheets(SheetName)

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If LastRow > 1 Then

        ws.Rows("2:" & LastRow).ClearContents

    End If

End Sub
