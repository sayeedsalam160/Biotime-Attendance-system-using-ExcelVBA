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
