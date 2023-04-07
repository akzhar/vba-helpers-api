Attribute VB_Name = "VbaHelper_GetCurrentUserEmail"
Option Explicit

Function GetCurrentUserEmail() As String
    ' Gets current user's email from Outlook
    
    Dim outlookObj As Object
    
    On Error Resume Next
    Set outlookObj = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    Dim namespaceObj As Object
    
    If outlookObj Is Nothing Then
        Set outlookObj = CreateObject("Outlook.Application")
        Set namespaceObj = outlookObj.GetNamespace("MAPI")
    End If

    Set outlookObj = Nothing: Set namespaceObj = Nothing
    GetCurrentUserEmail = outlookObj.Session.accounts.Item(1)
    
End Function