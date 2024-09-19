Attribute VB_Name = "VbaHelper_SetDropDownList"
Option Explicit

Function SetDropdown(ByRef rng As Range, ByVal src$)
    ' Sets dropdown list in the specified range

    ' @dependency: 69.bas
    ' @dependency: 133.bas
    Select Case True
        Case HasSubstring(src, GetListSeparator())
            src = src
        Case HasSubstring(src, "[") And HasSubstring(src, "]")
            src = "=INDIRECT(""" & src & """)"
        Case Else
            src = "=" & src
    End Select
    
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=src
    End With
End Function