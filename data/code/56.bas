Attribute VB_Name = "Helper56"
Option Explicit

Function SetDropDownList(ByRef rng As Range, ByVal source$)
    ' ф-ция устанавливает Data Validation с типом List
    rng.Validation.Delete
    rng.Validation.Add _
        Type:=xlValidateList, _
        Formula1:=IIf(InStr(1, source, ",", vbTextCompare), source, "=" & source)
End Function