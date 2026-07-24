Attribute VB_Name = "modUtilities"
Option Explicit

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
