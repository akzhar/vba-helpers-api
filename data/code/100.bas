Attribute VB_Name = "VbaHelper_GetDisplayResolution"
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long

Function GetDisplayResolution() As Long()
    ' Gets current display's resolution (X, Y)
    Dim arr(1) As Long
    arr(0) = GetSystemMetrics(0)
    arr(1) = GetSystemMetrics(1)
    GetDisplayResolution = arr

End Function