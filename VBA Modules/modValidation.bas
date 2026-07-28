Attribute VB_Name = "modValidation"
Option Explicit

Public Function ValidateHeaders() As Boolean

    Dim RequiredHeaders As Variant
    Dim i As Long

    RequiredHeaders = Array( _
        "Employee ID", _
        "First Name", _
        "Department", _
        "Date", _
        "Time", _
        "Punch State")

    For i = LBound(RequiredHeaders) To UBound(RequiredHeaders)

        If Not gHeaderMap.Exists(RequiredHeaders(i)) Then

            MsgBox "Missing Header: " & RequiredHeaders(i), vbCritical

            ValidateHeaders = False
            Exit Function

        End If

    Next i

    ValidateHeaders = True

End Function

Public Function IsValidEmployeeID(ByVal EmpID As Variant) As Boolean

    If Trim(EmpID & "") = "" Then
        IsValidEmployeeID = False
    Else
        IsValidEmployeeID = True
    End If

End Function

Public Function IsValidDateField(ByVal Value As Variant) As Boolean

    IsValidDateField = IsDate(Value)

End Function

Public Function IsValidTimeField(ByVal Value As Variant) As Boolean

    IsValidTimeField = IsDate(Value)

End Function
