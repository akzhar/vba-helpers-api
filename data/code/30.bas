Attribute VB_Name = "VbaHelper_AddBorders"
Option Explicit

Function AddBorders(ByRef rng As Range)
    ' Adds borders to range
    rng.Borders(xlEdgeLeft).LineStyle = xlContinuous
    rng.Borders(xlEdgeTop).LineStyle = xlContinuous
    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    rng.Borders(xlEdgeRight).LineStyle = xlContinuous
    rng.Borders(xlInsideVertical).LineStyle = xlContinuous
    rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Function