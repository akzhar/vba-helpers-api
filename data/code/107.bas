Attribute VB_Name = "Helper107"
Option Explicit

Function GetSelectedRows() As Variant()
    ' Gets array of rows from current selection (only visible rows)
        
    Dim selectedRows(), rngBounds() As String, rngRows(), isRange As Boolean, i&, j&, rowNum&
    Dim selectedRanges() As String: selectedRanges = Split(Replace(Selection.Address, "$", ""), ",")
    
    For i = LBound(selectedRanges) To UBound(selectedRanges)
        Erase rngBounds: rngBounds = Split(selectedRanges(i), ":")
        Erase rngRows: rngRows = GetRegExpMatches(selectedRanges(i), "\d+") ' @dependency: 60.bas
        isRange = False
        If UBound(rngBounds) > 0 Then
            isRange = InStr(1, selectedRanges(i), ":") And rngBounds(0) <> rngBounds(1)
        End If
        If isRange Then
            For j = CLng(rngRows(0)) To CLng(rngRows(1))
                rowNum = j
                If Not ActiveSheet.rows(rowNum).EntireRow.Hidden Then
                    Call AddToArr(selectedRows, rowNum) ' @dependency: 1.bas
                End If
            Next j
        Else
            rowNum = CLng(rngRows(0))
            Call AddToArr(selectedRows, rowNum) ' @dependency: 1.bas
        End If
    Next i
    GetSelectedRows = selectedRows
End Function