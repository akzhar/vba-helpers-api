Attribute VB_Name = "Helper80"
Option Explicit

' Порядковый номер листов
Const SHEET_1_NO& = 1
Const SHEET_2_NO& = 2

' Первая ячейка (левый верхний угол) с данными у таблицы на обеих листах
Const FIRST_ROW& = 2
Const FIRST_COL& = 1

' Цвета
Const COLOR_GREEN& = 5296274
Const COLOR_RED& = 255

Sub CompareTables()
    ' сверка 2-х одинаковых по структуре таблиц
    ' разница красится в красный, в комментарий проставляется diff сумма

    Call TurnUpdatesOn(False) ' @(id 51)

    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim iLastRow1&, iLastRow2&
    Dim iLastCol1&, iLastCol2&

    Set ws1 = ThisWorkbook.Sheets(SHEET_1_NO)
    Set ws2 = ThisWorkbook.Sheets(SHEET_2_NO)

    iLastRow2 = GetLastRow(ws2, FIRST_COL) ' @(id 64)
    iLastCol2 = GetLastColumn(ws2, FIRST_ROW - 1) ' @(id 65)

    iLastRow1 = GetLastRow(ws1, FIRST_COL) ' @(id 64)
    iLastCol1 = GetLastColumn(ws1, FIRST_ROW - 1) ' @(id 65)

    ws2.Range(ws2.Cells(FIRST_ROW, FIRST_COL), ws2.Cells(iLastRow2, FIRST_COL)).Interior.Color = COLOR_RED

    ws1.Range(ws1.Cells(FIRST_ROW, FIRST_COL), ws1.Cells(iLastRow1, FIRST_COL)).Interior.Color = COLOR_RED

    If iLastCol2 <> iLastCol1 Then
        Call Ended
        MsgBox "Количество и порядок следования столбцов в 2-х сравниваемых таблицах должны совпадать", vbCritical
        Exit Sub
    End If

    Dim i&, j&, k&
    Dim cellUnique1 As Range, cellUnique2 As Range
    Dim cell1 As Range, cell2 As Range

    For i = FIRST_ROW To iLastRow1

        For j = FIRST_ROW To iLastRow2
            
            Set cellUnique1 = ws1.Cells(i, FIRST_COL)
            Set cellUnique2 = ws2.Cells(j, FIRST_COL)
            
            If cellUnique1.Value = cellUnique2.Value Then

                cellUnique1.Interior.Color = COLOR_GREEN
                cellUnique2.Interior.Color = COLOR_GREEN
                
                For k = FIRST_COL + 1 To iLastCol1 ' или iLastCol2

                    Set cell1 = ws1.Cells(i, k)
                    Set cell2 = ws2.Cells(j, k)
                    
                    If cell1.Value = cell2.Value Then

                        cell1.Interior.Color = COLOR_GREEN
                        cell2.Interior.Color = COLOR_GREEN

                    Else

                        cell1.Interior.Color = COLOR_RED
                        cell2.Interior.Color = COLOR_RED

                        If IsNumeric(cell1.Value) And IsNumeric(cell2.Value) Then
                            cell1.AddComment "Разница = " & cell1.Value - cell2.Value
                            cell2.AddComment "Разница = " & cell2.Value - cell1.Value
                        End If

                    End If

                Next k

            End If
            
        Next j

    Next i

    MsgBox "Сверка 2-х таблиц завершена", vbInformation

    Call TurnUpdatesOn(True) ' @(id 51)

End Sub

Sub ResetTables()
    ' возврат обеих таблиц в исходное состояние

    Call TurnUpdatesOn(False) ' @(id 51)

    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim iLastRow1&, iLastRow2&
    Dim iLastCol1&, iLastCol2&

    Set ws1 = ThisWorkbook.Sheets(SHEET_1_NO)
    Set ws2 = ThisWorkbook.Sheets(SHEET_2_NO)

    iLastRow2 = GetLastRow(ws2, FIRST_COL) ' @(id 64)
    iLastCol2 = GetLastColumn(ws2, FIRST_ROW - 1) ' @(id 65)

    With ws2.Range(ws2.Cells(FIRST_ROW, FIRST_COL), ws2.Cells(iLastRow2, iLastCol2))
      .Interior.Color = xlNone
      .ClearComments
    End with
    Application.Goto [A1], True

    iLastRow1 = GetLastRow(ws1, FIRST_COL) ' @(id 64)
    iLastCol1 = GetLastColumn(ws1, FIRST_ROW - 1) ' @(id 65)

    With ws1.Range(ws1.Cells(FIRST_ROW, FIRST_COL), ws1.Cells(iLastRow1, iLastCol1))
      .Interior.Color = xlNone
      .ClearComments
    End With
    Application.Goto [A1], True

    Call TurnUpdatesOn(True) ' @(id 51)

End Sub