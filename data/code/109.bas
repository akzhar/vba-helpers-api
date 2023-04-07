Attribute VB_Name = "VbaHelper_OpenUrl"
Option Explicit

Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal Operation As String, _
    ByVal Filename As String, _
    Optional ByVal Parameters As String, _
    Optional ByVal Directory As String, _
    Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Function OpenUrl(ByVal url$)
    Call ShellExecute(0, "Open", url)
End Function