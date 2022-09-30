Attribute VB_Name = "Helper45"
Option Explicit

Function GetCurrentUserEmail() As String
    ' Gets current user's email from Outlook
    
    Dim objOutlook As Object
    
    On Error Resume Next
    Set objOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    Dim objNameSpace As Object
    
    If objOutlook Is Nothing Then
        Set objOutlook = CreateObject("Outlook.Application")
        Set objNameSpace = objOutlook.GetNamespace("MAPI")
    End If
        
    GetCurrentUserEmail = objOutlook.Session.accounts.Item(1)
    
End Function