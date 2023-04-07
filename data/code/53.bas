Attribute VB_Name = "VbaHelper_HasValidation"
Option Explicit

Function HasValidation(ByRef rng As Range) As Boolean
    ' Checks if range has Data Validation set in it

    Dim validatioType: validatioType = Null
    On Error Resume Next
    validatioType = rng.Validation.Type
    On Error GoTo 0
    HasValidation = Not IsNull(validatioType)
End Function