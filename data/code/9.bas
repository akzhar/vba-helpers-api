Attribute VB_Name = "Helper9"
Option Explicit

Function SortArr(ByRef arr(), ByVal N&, Optional ByVal isDesc As Boolean = True) As Variant
    ' ф-ция сортирует переданный 2 мерный массив в столбце N по алфавиту
    ' в столбце N должен быть текст
    ' desc: по убыванию, от большего к меньшему (<)
    ' asc: по возрастанию, от меньшего к большему (>)

    If N > UBound(arr, 1) Or N < LBound(arr, 1) Then
       MsgBox "Нет такого столбца в массиве", vbExclamation
       Exit Function
    End if

    Dim check As Boolean, iCount&, jCount&, nCount&, condition As Boolean, tempArr()

    ReDim tempArr(UBound(arr, 2)) As Variant

    Do Until check
        check = True
        For iCount = LBound(arr, 2) To UBound(arr, 2) - 1
            condition = Iif( _
              isDesc, _
              Left(arr(N, iCount), 1) < Left(arr(N, iCount + 1), 1), _
              Left(arr(N, iCount), 1) > Left(arr(N, iCount + 1), 1) _
            )
            If condition Then
                For jCount = LBound(arr, 1) To UBound(arr, 1)
                    tempArr(jCount) = arr(jCount, iCount)
                    arr(jCount, iCount) = arr(jCount, iCount + 1)
                    arr(jCount, iCount + 1) = tempArr(jCount)
                    check = False
                Next
            End If
        Next
    Loop

    SortArr = arr

End Function