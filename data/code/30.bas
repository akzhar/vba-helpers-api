Attribute VB_Name = "VbaHelper_AddBorders"
Option Explicit

Function AddBorders(ByRef rng As Range, Optional ByVal lineStyle& = xlContinuous, Optional ByVal lineColor& = vbBlack)
    ' Adds borders to range
    ' linesStyle is XlLineStyle enum https://learn.microsoft.com/ru-ru/office/vba/api/excel.xllinestyle
    With rng.Borders(xlEdgeLeft)
      .LineStyle = lineStyle
      .Color = lineColor
    End With
    With rng.Borders(xlEdgeTop)
      .LineStyle = lineStyle
      .Color = lineColor
    End With
    With rng.Borders(xlEdgeBottom)
      .LineStyle = lineStyle
      .Color = lineColor
    End With
    With rng.Borders(xlEdgeRight)
      .LineStyle = lineStyle
      .Color = lineColor
    End With
    With rng.Borders(xlInsideVertical)
      .LineStyle = lineStyle
      .Color = lineColor
    End With
    With rng.Borders(xlInsideHorizontal)
      .LineStyle = lineStyle
      .Color = lineColor
    End With
End Function