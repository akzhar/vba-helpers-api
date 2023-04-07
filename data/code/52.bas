Attribute VB_Name = "VbaHelper_AdjustNotes"
Option Explicit

Function AdjustNotes(ByRef ws As Worksheet)
    ' Adjusts the size of all cell notes on the worksheet to the size of their contents
    Dim com As Comment
    For Each com In ws.Comments
        com.Shape.TextFrame.AutoSize = True
    Next com
End Function