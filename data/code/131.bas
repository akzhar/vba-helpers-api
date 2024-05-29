Attribute VB_Name = "VbaHelper_GetQuarterByDate"
Option Explicit

Function GetQuarterByDate(ByVal d As Date, Optional asText As Boolean = False) As Variant
    ' Returns number of quarter by date
    
    Dim monthNum&: monthNum = Month(d)
    Dim quarterNum&: quarterNum = -1
    
    GetQuarterByDate = quarterNum

    Select Case monthNum
        Case 1 To 3
            quarterNum = 1
        Case 4 To 6
            quarterNum = 2
        Case 7 To 9
            quarterNum = 3
        Case 10 To 12
            quarterNum = 4
    End Select

    If quarterNum = -1 Then Exit Function

    Dim o As Object: Set o = CreateObject("Scripting.Dictionary")
    o.Add 1, "I": o.Add 2, "II": o.Add 3, "III": o.Add 4, "IV"
    
    GetQuarterByDate = Iif(asText, o(quarterNum), quarterNum)

End Function