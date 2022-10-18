VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private startTime&
Private finishTime&
Private executionTime$

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub Start()
    startTime = 0
    finishTime = 0
    executionTime = 0
    startTime = Timer()
End Sub

Public Sub Finish()
    If startTime <> 0 Then
        finishTime = Timer()
        executionTime = Format((finishTime - startTime) / 86400, "hh:mm:ss")
    End If
End Sub

Public Function GetTime() As String
    If finishTime = 0 Then Call Finish
    GetTime = executionTime
End Function
