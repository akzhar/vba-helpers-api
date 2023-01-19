Attribute VB_Name = "Helper104"
Option Explicit

Function GetCurrentUserLogin() As String
    ' Gets current user's login
    
    GetCurrentUserLogin = Environ("UserName")
    
End Function