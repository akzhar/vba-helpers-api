Attribute VB_Name = "VbaHelper_IsNamedRangeExists"
Option Explicit

Function IsNamedRangeExists(ByRef ws As Worksheet, ByVal rngName$) As Boolean
    ' Checks if named range exists or not

    Dim rng As Range

    On Error Resume Next
    Set rng = ws.Range(rngName)

    IsNamedRangeExists = CBool(Err.Number = 0)

End Function