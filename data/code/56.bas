Attribute VB_Name = "Helper56"
Option Explicit

Function SetDropDownList(ByRef rng As Range, ByVal listName$)
    ' ф-ция устанавливает для rng Data Validation с типом List (source = listName)
   With rng.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=" & listName
    End With
End Function