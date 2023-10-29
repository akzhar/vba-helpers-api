Attribute VB_Name = "VbaHelper_RemoveCondFormatting"
Option Explicit

Function RemoveCondFormatting(ByRef rng As Range)
    ' Removes condition formatting from rng
    If rng.FormatConditions.Count Then
        Dim fc As FormatCondition
        For Each fc In rng.FormatConditions
            fc.Delete
        Next fc
    End If    
End Function