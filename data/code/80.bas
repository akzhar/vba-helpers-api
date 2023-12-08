Attribute VB_Name = "VbaHelper_CompareTables"
Option Explicit

' Worksheets number
Private Const SHEET_1_NAME$ = "Table 1"
Private Const SHEET_2_NAME$ = "Table 2"

' Left top cornerof the table on both worksheets
Private Const FIRST_ROW& = 2
Private Const FIRST_COL& = 1

' Colors
Private Const COLOR_GREEN& = 5296274
Private Const COLOR_RED& = 255

Sub CompareTables()
    ' Comparison of 2 tables with the same structure (identical headers, quantity and order of columns)

    Call TurnUpdatesOn(False) ' @dependency: 51.bas

    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim iLastRow1&, iLastRow2&
    Dim iLastCol1&, iLastCol2&

    Set ws1 = ThisWorkbook.Sheets(SHEET_1_NAME)
    Set ws2 = ThisWorkbook.Sheets(SHEET_2_NAME)

    iLastRow2 = GetLastRow(ws2, FIRST_COL) ' @dependency: 64.bas
    iLastCol2 = GetLastColumn(ws2, FIRST_ROW - 1) ' @dependency: 65.bas

    iLastRow1 = GetLastRow(ws1, FIRST_COL) ' @dependency: 54.bas
    iLastCol1 = GetLastColumn(ws1, FIRST_ROW - 1) ' @dependency: 65.bas

    ws2.Range(ws2.Cells(FIRST_ROW, FIRST_COL), ws2.Cells(iLastRow2, FIRST_COL)).Interior.Color = COLOR_RED

    ws1.Range(ws1.Cells(FIRST_ROW, FIRST_COL), ws1.Cells(iLastRow1, FIRST_COL)).Interior.Color = COLOR_RED

    If iLastCol2 <> iLastCol1 Then
        Call TurnUpdatesOn(True) ' @dependency: 51.bas
        MsgBox "The number and order of columns in the 2 compared tables must match", vbCritical
        Exit Sub
    End If

    Dim i&, j&, k&
    Dim cellUnique1 As Range, cellUnique2 As Range
    Dim cell1 As Range, cell2 As Range

    For i = FIRST_ROW To iLastRow1

        For j = FIRST_ROW To iLastRow2
            
            Set cellUnique1 = ws1.Cells(i, FIRST_COL)
            Set cellUnique2 = ws2.Cells(j, FIRST_COL)
            
            If CStr(cell1.Value) = CStr(cell2.Value) Then

                cellUnique1.Interior.Color = COLOR_GREEN
                cellUnique2.Interior.Color = COLOR_GREEN
                
                For k = FIRST_COL + 1 To iLastCol1 ' or iLastCol2

                    Set cell1 = ws1.Cells(i, k)
                    Set cell2 = ws2.Cells(j, k)
                    
                    If cell1.Value = cell2.Value Then

                        cell1.Interior.Color = COLOR_GREEN
                        cell2.Interior.Color = COLOR_GREEN

                    Else

                        cell1.Interior.Color = COLOR_RED
                        cell2.Interior.Color = COLOR_RED

                        If IsNumeric(cell1.Value) And IsNumeric(cell2.Value) Then
                            cell1.AddComment "Difference = " & CStr(cell1.Value - cell2.Value)
                            cell2.AddComment "Difference = " & CStr(cell2.Value - cell1.Value)
                        End If

                    End If

                Next k

            End If
            
        Next j

    Next i

    MsgBox "Comparison of 2 tables completed", vbInformation

    Call TurnUpdatesOn(True) ' @dependency: 51.bas

End Sub

Sub ResetTables()
    ' Returns both tables to their original state

    Call TurnUpdatesOn(False) ' @dependency: 51.bas

    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim iLastRow1&, iLastRow2&
    Dim iLastCol1&, iLastCol2&

    Set ws1 = ThisWorkbook.Sheets(SHEET_1_NAME)
    Set ws2 = ThisWorkbook.Sheets(SHEET_2_NAME)

    iLastRow2 = GetLastRow(ws2, FIRST_COL) ' @dependency: 64.bas
    iLastCol2 = GetLastColumn(ws2, FIRST_ROW - 1) ' @dependency: 65.bas

    With ws2.Range(ws2.Cells(FIRST_ROW, FIRST_COL), ws2.Cells(iLastRow2, iLastCol2))
      .Interior.Color = xlNone
      .ClearComments
    End with
    Application.Goto [A1], True

    iLastRow1 = GetLastRow(ws1, FIRST_COL) ' @dependency: 64.bas
    iLastCol1 = GetLastColumn(ws1, FIRST_ROW - 1) ' @dependency: 65.bas

    With ws1.Range(ws1.Cells(FIRST_ROW, FIRST_COL), ws1.Cells(iLastRow1, iLastCol1))
      .Interior.Color = xlNone
      .ClearComments
    End With
    Application.Goto [A1], True

    Call TurnUpdatesOn(True) ' @dependency: 51.bas

End Sub