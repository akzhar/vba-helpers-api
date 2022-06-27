Attribute VB_Name = "Helper53"
Option Explicit

Function HasValidation(ByRef rng As Range) As Boolean
    ' ф-ция проверяет наличие Data  Validation в ячейке
    Dim validatioType: validatioType = Null

    On Error Resume Next
    validatioType = rng.Validation.Type
    On Error GoTo 0

    HasValidation = Not IsNull(validatioType)
End Function