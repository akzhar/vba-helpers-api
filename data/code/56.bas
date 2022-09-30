Attribute VB_Name = "Helper56"
Option Explicit

Function SetDropDownList(ByRef rng As Range, ByVal source$)
    ' Sets dropdown list in the specified range

    Select Case True
        Case Includes(source, ",") ' @(id 69)
            source = source
        Case Includes(source, "[") And Includes(source, "]") ' @(id 69)
            source = "=INDIRECT(""" & source & """)"
        Case Else
            source = "=" & source
    End Select
    
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:=source
    End With
End Function