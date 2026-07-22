Attribute VB_Name = "modImport"
Option Explicit

Public Function SelectACSFile() As Boolean

    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

        .Title = "Select ACS Attendance Report"

        .AllowMultiSelect = False

        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls"

        If .Show <> -1 Then
            SelectACSFile = False
            Exit Function
        End If

        gACSFilePath = .SelectedItems(1)

    End With

    SelectACSFile = True

End Function

Public Sub ImportACSData()

    Dim wbACS As Workbook
    Dim wsACS As Worksheet
    Dim wsRaw As Worksheet

    Dim LastRow As Long
    Dim LastCol As Long

    Application.ScreenUpdating = False

    Set wsRaw = ThisWorkbook.Worksheets("Raw Import")

    Set wbACS = Workbooks.Open(gACSFilePath)

    Set wsACS = wbACS.Sheets(1)

    LastRow = wsACS.Cells(wsACS.Rows.Count, "A").End(xlUp).Row
    LastCol = wsACS.Cells(2, wsACS.Columns.Count).End(xlToLeft).Column

    wsRaw.Cells.Clear

    'Copy headers (ACS Row 2 -> Raw Row 1)
    wsACS.Range(wsACS.Cells(2, 1), wsACS.Cells(2, LastCol)).Copy
    wsRaw.Range("A1").PasteSpecial xlPasteValues

    'Copy data (ACS Row 3+ -> Raw Row 2+)
    wsACS.Range(wsACS.Cells(3, 1), wsACS.Cells(LastRow, LastCol)).Copy
    wsRaw.Range("A2").PasteSpecial xlPasteValues

    wbACS.Close False

    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    MsgBox LastRow - 2 & " records imported successfully.", vbInformation

End Sub
