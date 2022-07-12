Attribute VB_Name = "Helper76"
Option Explicit

Function CONCATENATEIF(ByRef rngToCheck As Range, ByRef rngToConcat As Range, ByVal pattern$, Optional separator$ = " ") As String
    ' аналог встроенной ф-ции CONCATENATE, но с возможностью задать условие конкатенации
    '---------------------------------------------------------------------------------------
    ' rngToCheck - диапазон с критериями (указывается один столбец)
    ' rngToConcat - из этого диапазона берется значение для сцепления
    ' pattern - критерий (шаблон, с которым будет происходить сверка значений из rngToCheck)
    ' (? - любой отдельный знак, * - ноль или более символов, # - любая отдельная цифра)
    ' separator - разделитель, по умолчанию пробел
    ' условие конкатенации: если значение в ячейке из rngToCheck соответствет pattern
    '---------------------------------------------------------------------------------------
    Application.Volatile True
    Dim cell As Range, str$
    'Set rngToCheck = Intersect(rngToCheck, ActiveSheet.UsedRange)
    'Set rngToConcat = Intersect(rngToConcat, ActiveSheet.UsedRange)
    For Each cell In rngToCheck
       If cell.Value Like pattern And Trim(rngToConcat.Cells(cell.Row - rngToCheck.Row + 1, 1)) <> "" Then
          str = str & IIf(str <> "", separator, "") & rngToConcat.Cells(cell.Row - rngToCheck.Row + 1, 1)
       End If
    Next cell
    CONCATENATEIF = str
End Function