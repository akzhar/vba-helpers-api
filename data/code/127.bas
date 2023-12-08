Attribute VB_Name = "VbaHelper_SelectRng"
Option Explicit

Function SelectRng(ByVal titleMessage$) As Range
    ' Allows to select a range from the sheet
    On Error Resume Next
    Set SelectRng = Application.InputBox(titleMessage, "Range picker", Selection.Address, Type:=8)
    On Error GoTo 0
    If Not SelectRng Is Nothing Then
        SelectRng.Select
    End If
End Function