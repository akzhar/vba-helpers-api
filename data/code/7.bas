Attribute VB_Name = "VbaHelper_Filter2DArr"
Option Explicit

Function Filter2DArr(ByRef arr(), ParamArray args()) As Variant()
    ' Filters 2-dim array using an arbitrary number of filtering criteria
    Filter2DArr = Array()
    
    On Error Resume Next

    If UBound(args) = -1 Then
        Debug.Print "Error: filters required"
        Exit Function
    End If

    ReDim Filters(0 To UBound(args) + 1, 1 To 2)
    Err.Clear

    Dim i&: i = UBound(arr, 2)

    If Err.Number > 0 Then
        Debug.Print "Error: two dimensional array required"
        Exit Function
    End If

    Dim arg$, mask$, col&
    Dim filtersCount&
    Dim filtersArr(): ReDim filtersArr(UBound(args), 1)

    ' Get all the criteria
    For i = LBound(args) To UBound(args)
        arg = args(i)
        If Not IsMissing(arg) Then
            If arg Like "#*=*" Then
                filtersCount = filtersCount + 1
                col = Val(Split(arg, "=")(0)) ' column
                mask = Split(arg, "=", 2)(1) ' mask
                filtersArr(i, 0) = col
                filtersArr(i, 1) = mask
            Else
                Debug.Print "Error: invalid filter '" & arg & "'"
            End If
        End If
    Next i

    If filtersCount = 0 Then
        Debug.Print "Error: all filters are empty"
        Exit Function
    End If

    Dim rowsCount&, j&
    Dim checksArr() As Boolean: ReDim checksArr(LBound(arr, 1) To UBound(arr, 1))

    ' check all the rows
    For i = LBound(arr, 1) To UBound(arr, 1)
        checksArr(i) = True
        ' check all the criteria for each row
        For j = 1 To filtersCount
            col = filtersArr(j - 1, 0)
            mask = filtersArr(j - 1, 1)
            If Not (arr(i, col) Like mask) Then
                checksArr(i) = False
                Exit For
          End If
      Next j
      rowsCount = rowsCount - checksArr(i)
    Next i

    If rowsCount = 0 Then
        Debug.Print "There are no rows matched all the filtering creteria"
        Exit Function
    End If

    Dim rowNum&
    ReDim filteredArr(0 To rowsCount - 1, LBound(arr, 2) To UBound(arr, 2))

    ' collect all the rows which passed filtering
    For i = LBound(arr, 1) To UBound(arr, 1)
        If checksArr(i) Then
            rowNum = rowNum + 1
            For j = LBound(arr, 2) To UBound(arr, 2)
                filteredArr(rowNum - 1, j) = arr(i, j)
            Next j
        End If
    Next i

    On Error GoTo 0

    Filter2DArr = filteredArr
End Function