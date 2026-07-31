Attribute VB_Name = "modUtilities"
Public Sub BuildHeaderMap()

    Dim ws As Worksheet
    Dim LastCol As Long
    Dim Col As Long
    Dim HeaderName As String

    Set ws = ThisWorkbook.Worksheets("Raw Import")

    Set gHeaderMap = CreateObject("Scripting.Dictionary")

    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For Col = 1 To LastCol

        HeaderName = Trim(ws.Cells(1, Col).Value)

        If HeaderName <> "" Then

            If Not gHeaderMap.Exists(HeaderName) Then
                gHeaderMap.Add HeaderName, Col
            End If

        End If

    Next Col

End Sub

Public Function GetAttendanceDate( _
    ByVal PunchDate As Date, _
    ByVal PunchTime As Date) As Date

    Dim TimeOnly As Date

    TimeOnly = TimeValue(PunchTime)

    ' 12:00 AM to 3:00 AM belongs to previous day
    If TimeOnly <= TimeSerial(3, 0, 0) Then

        GetAttendanceDate = PunchDate - 1

    Else

        GetAttendanceDate = PunchDate

    End If

End Function

Public Function CalculateHours( _
    ByVal PunchIn As Date, _
    ByVal PunchOut As Date) As Double

    CalculateHours = Round((PunchOut - PunchIn) * 24, 2)

End Function

Public Function GetAttendanceStatus( _
    ByVal PunchIn As Variant, _
    ByVal PunchOut As Variant) As String

    If Not IsDate(PunchIn) Or Not IsDate(PunchOut) Then

        GetAttendanceStatus = "Incomplete"

    Else

        GetAttendanceStatus = "Present"

    End If

End Function

