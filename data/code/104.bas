Attribute VB_Name = "VbaHelper_GetCurrentUserLogin"
Option Explicit

Function GetCurrentUserLogin() As String
    ' Gets current user's login
    GetCurrentUserLogin = Environ("UserName")
End Function