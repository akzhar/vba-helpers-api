Attribute VB_Name = "Helper52"
Option Explicit

Function FitComments(ByRef ws as Worksheet)
    ' ф-ция ресайзит комментарии на листе
    
    Dim com As Comment

    For Each com In ws.Comments
        com.Shape.TextFrame.AutoSize = True
    Next com
    
End Function