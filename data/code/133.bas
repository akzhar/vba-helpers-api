Attribute VB_Name = "VbaHelper_GetListSeparator"
Option Explicit

Function GetListSeparator() As String
    ' Gets list separator symbol from Windows register
    Dim regObj As Object: Set regObj = CreateObject("WScript.Shell")
    Dim regKey$: regKey = regObj.RegRead("HKEY_CURRENT_USER\Control Panel\International\Slist")
    If regKey = "" Then
        'MsgBox "There is no registry value for this key", vbExclamation
        GetListSeparator = ""
    Else
        GetListSeparator = regKey
    End If
    Set regObj = Nothing
End Function
