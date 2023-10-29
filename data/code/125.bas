Attribute VB_Name = "VbaHelper_SetCondFormatting"
Option Explicit

Function SetCondFormatting(ByRef rng As Range, ByVal compareOperator$, ByVal criteriaValue As Double, Optional isRed As Boolean = False)
    ' Set red / green condition formating
    Dim xlOperator
    Select Case compareOperator
      Case ">":
        xlOperator = xlGreater
      Case "<":
        xlOperator = xlLess
      Case ">=":
        xlOperator = xlGreaterEqual
      Case "<=":
        xlOperator = xlLessEqual
      Case "=":
        xlOperator = xlEqual
      Case Else:
        End Function
    End Select
    With rng
        Call RemoveCondFormatting(rng) ' @dependency: 124.bas
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlOperator, Formula1:="=" & CStr(criteriaValue)
        .FormatConditions(1).Font.Color = IIf(isRed, RGB(156, 0, 6), RGB(0, 97, 0)) ' red / green
        .FormatConditions(1).Interior.Color = IIf(isRed, RGB(255, 199, 206), RGB(198, 239, 206)) ' red / green
    End With
End Function