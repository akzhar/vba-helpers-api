VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class to manipulate 1-dim array

Private internalArr()
    
Private Sub Class_Initialize()
    internalArr = Array()
End Sub

Private Sub Class_Terminate()
    Call Reset
End Sub

' Properties

Public Property Get Values() As Variant

    Values = internalArr
    
End Property

Public Property Let Values(items As Variant) ' https://stackoverflow.com/a/46841174
    
    internalArr = items
    
End Property

' Private methods

Private Function IsInitialised(ByRef optArr() As Variant) As Boolean
    
    On Error Resume Next
    IsInitialised = IsNumeric(UBound(optArr))
    On Error GoTo 0
    
End Function

' Public methods

Public Sub Reset()

    Erase internalArr
    
End Sub

Public Function GetLenght(Optional ByRef optArr As Variant) As Long
    
    Dim arr(): arr = IIf(IsMissing(optArr), internalArr, optArr)
    
    If IsEmpty(arr) Or Not IsInitialised(arr) Then
        GetLenght = 0
    Else
        GetLenght = UBound(arr) - LBound(arr) + 1
    End If
    
    Erase arr
    
End Function

' [...x,y] <- y
Public Sub Push(ByVal element As Variant, Optional ByRef optArr As Variant)
    
    Dim arr(): arr = IIf(IsMissing(optArr), internalArr, optArr)
    
    If GetLenght(optArr:=arr) = 0 Then
        ReDim arr(0)
    Else
        ReDim Preserve arr(UBound(arr) + 1)
    End If
    
    If TypeName(element) = "Dictionary" Then
        Set arr(UBound(arr)) = element
    Else
        arr(UBound(arr)) = element
    End If
    
    If IsMissing(optArr) Then
        internalArr = arr
    Else
        optArr = arr
    End If
    
    Erase arr
    
End Sub

' [...x,y] -> y
Public Function Pop() As Variant
    
    If GetLenght() = 0 Then
        Pop = Null
    Else
        Pop = internalArr(UBound(internalArr))
        If UBound(internalArr) = 0 Then
            Call Reset
        Else
            ReDim Preserve internalArr(UBound(internalArr) - 1)
        End If
    End If
    
End Function

' x <- [x,y...]
Public Function Shift() As Variant

    Dim arr(): arr = Array()
    
    If GetLenght() = 0 Then
        Shift = Null
    Else
        Shift = internalArr(0)
    
        Dim i&
        For i = LBound(internalArr) + 1 To UBound(internalArr)
            Push element:=internalArr(i), optArr:=arr
        Next i
    End If
    
    internalArr = arr
    Erase arr
    
End Function

' x -> [x,y...]
Public Sub Unshift(ByVal element As Variant)
    
    If GetLenght() = 0 Then
        Push element:=element
        Exit Sub
    End If
    
    Dim arr(): arr = Array()
    
    Push element:=element, optArr:=arr
    
    Dim i&
    For i = LBound(internalArr) To UBound(internalArr)
        Push element:=internalArr(i), optArr:=arr
    Next i
    
    internalArr = arr
    Erase arr
    
End Sub

Public Sub PushFromRng(ByRef rng As Range)

    Dim cell As Range
    For Each cell In rng
        Push element:=cell.value
    Next cell
    
End Sub

Public Sub PasteToRng(ByRef rng As Range, Optional ByVal inRow As Boolean = False)
    
    If GetLenght() = 0 Then Exit Sub
    
    If inRow Then
        rng.Resize(1, UBound(internalArr) + 1).value = internalArr
    Else
        rng.Resize(UBound(internalArr) + 1, 1).value = Application.WorksheetFunction.Transpose(internalArr)
    End If
    
End Sub

Public Function Includes(ByVal element As Variant, Optional ByRef optArr As Variant) As Boolean
    
    Dim arr(): arr = IIf(IsMissing(optArr), internalArr, optArr)

    Includes = False
    
    If GetLenght(optArr:=arr) = 0 Then Exit Function
    
    Dim i&
    For i = LBound(arr) To UBound(arr)
        On Error GoTo NextItem
        If arr(i) = element Then
            Includes = True
            Exit For
        End If
NextItem:
        On Error GoTo 0
    Next i
    
    Erase arr

End Function

Public Sub RemoveDuplicates()
    
    Dim arr(): arr = Array()
    Dim isDuplicate As Boolean
    
    Dim i&
    For i = LBound(internalArr) To UBound(internalArr)
        isDuplicate = Includes(element:=internalArr(i), optArr:=arr)
        If Not isDuplicate Then
            Push element:=internalArr(i), optArr:=arr
        End If
    Next i
    
    internalArr = arr
    Erase arr

End Sub

Public Function Combine(ByRef arr As Variant) As Variant()
    
    Dim originalArr(): originalArr = internalArr
    
    Dim i&
    For i = LBound(arr) To UBound(arr)
        Push element:=arr(i), optArr:=originalArr
    Next i
    
    Combine = originalArr
    
End Function

Public Sub Sort(Optional ByVal isDesc As Boolean = False)

    Dim i&, j&, condition As Boolean, temp
    
    For i = LBound(internalArr) To UBound(internalArr) - 1
        For j = i + 1 To UBound(internalArr)
            condition = IIf(isDesc, internalArr(i) < internalArr(j), internalArr(i) > internalArr(j))
            If condition Then
                temp = internalArr(j)
                internalArr(j) = internalArr(i)
                internalArr(i) = temp
            End If
        Next j
    Next i

End Sub

Public Function Filter(ByVal filterValue As Variant) As Variant()

    Dim arr(): arr = Array()
    
    Dim i&
    For i = LBound(internalArr) To UBound(internalArr)
        If filterValue = internalArr(i) Then
            Push element:=internalArr(i), optArr:=arr
        End If
    Next i

    Filter = arr
    
End Function

Public Function Filter2(ByVal assertFunc$) As Variant()

    Dim arr(): arr = Array()
    
    Dim i&
    For i = LBound(internalArr) To UBound(internalArr)
        If Application.Run(assertFunc, internalArr(i)) Then
            Push element:=internalArr(i), optArr:=arr
        End If
    Next i

    Filter2 = arr
    
End Function

Public Function IndexOf(ByVal element As Variant) As Long
    
    IndexOf = -1
    
    Dim i&
    For i = LBound(internalArr) To UBound(internalArr)
        If element = internalArr(i) Then
            IndexOf = i
            Exit For
        End If
    Next i

End Function

Public Function ToString() As String
    
    ToString = "[ " & Join(internalArr, ", ") & " ]"

End Function
