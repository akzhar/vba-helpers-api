Attribute VB_Name = "VbaHelper_SetDropDownList"
Option Explicit

Function SetDropdown(ByRef rng As Range, ByVal src$)
    ' Sets dropdown list in the specified range

    Select Case True
        Case HasSubstring(src, ",") ' @dependency: 69.bas
            src = src
        Case HasSubstring(src, "[") And HasSubstring(src, "]") ' @dependency: 69.bas
            src = "=INDIRECT(""" & src & """)"
        Case Else
            src = "=" & src
    End Select
    
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=src
    End With
End Function