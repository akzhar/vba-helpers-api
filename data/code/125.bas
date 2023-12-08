Attribute VB_Name = "VbaHelper_SetCondFormatting"
Option Explicit

Function SetCondFormatting(ByRef rng As Range, ByVal compareOperator$, ByVal limitValue As Double, Optional isRed As Boolean = False)
    ' Set red / green condition formating
    ' compareOperator is XlFormatConditionOperator enum
    ' https://learn.microsoft.com/ru-ru/office/vba/api/excel.xlformatconditionoperator
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
        Exit Function
    End Select

    Dim fontColor&: fontColor = IIf(isRed, RGB(156, 0, 6), RGB(0, 97, 0)) ' red / green
    Dim backColor&: backColor = IIf(isRed, RGB(255, 199, 206), RGB(198, 239, 206)) ' red / green

    If rng.FormatConditions.Count Then
        Dim fc As FormatCondition
        For Each fc In rng.FormatConditions
          If fc.Font.Color = fontColor And fc.Interior.Color = backColor Then
            fc.Delete
          End If
        Next fc
    End If

    With rng
        .FormatConditions.Add _
          Type:=xlCellValue, _
          Operator:=xlOperator, _
          Formula1:="=" & CStr(limitValue)
        .FormatConditions(.FormatConditions.Count).Font.Color = fontColor
        .FormatConditions(.FormatConditions.Count).Interior.Color = backColor
    End With
End Function