Attribute VB_Name = "Helper30"
Option Explicit

Function AddBorders(ByRef rng As Range)
    ' ф-ция добавляет границы у диапазона rng
    rng.Borders(xlEdgeLeft).LineStyle = xlContinuous
    rng.Borders(xlEdgeTop).LineStyle = xlContinuous
    rng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    rng.Borders(xlEdgeRight).LineStyle = xlContinuous
    rng.Borders(xlInsideVertical).LineStyle = xlContinuous
    rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Function