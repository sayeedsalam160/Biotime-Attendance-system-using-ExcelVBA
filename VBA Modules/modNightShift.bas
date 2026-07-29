Attribute VB_Name = "modNightShift"
Public Function IsNightIn(ByVal PunchTime As Date) As Boolean

    Dim T As Date

    T = TimeValue(PunchTime)

    IsNightIn = (T >= TimeValue("00:00:00") And _
                 T <= TimeValue("03:00:00"))

End Function

Public Function IsNightOut(ByVal PunchTime As Date) As Boolean

    Dim T As Date

    T = TimeValue(PunchTime)

    IsNightOut = (T >= TimeValue("06:00:00") And _
                  T <= TimeValue("12:00:00"))

End Function

