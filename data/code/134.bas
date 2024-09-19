Attribute VB_Name = "VbaHelper_IsDropdown"
Option Explicit

Function IsDropdown(ByRef rng As Range) As Boolean
    ' Checks if specified cell contains a native drop-down list (data validation with type = list)
    IsDropdown = False
    If rng.Count = 1 And HasValidation(rng) Then '@dependency: 53.bas
        If rng.Validation.Type = 3 Then
            IsDropdown = True
        End If
    End If
End Function
