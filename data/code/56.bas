Attribute VB_Name = "VbaHelper_SetDropDownList"
Option Explicit

Function SetDropDownList(ByRef rng As Range, ByVal source$)
    ' Sets dropdown list in the specified range

    Select Case True
        Case HasSubstring(source, ",") ' @dependency: 69.bas
            source = source
        Case HasSubstring(source, "[") And HasSubstring(source, "]") ' @dependency: 69.bas
            source = "=INDIRECT(""" & source & """)"
        Case Else
            source = "=" & source
    End Select
    
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=source
    End With
End Function