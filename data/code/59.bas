Attribute VB_Name = "Helper59"
Option Explicit

Function ProtectWs(ByRef ws As Worksheet, ByVal password$)
    ' ф-ция ставит пароль на лист
    With ws
        .Protect Password:=password, _
         AllowFiltering:=True ' автофильтр вкл
        .EnableSelection = xlNoRestrictions
    End With
End Function

Function UnProtectWs(ByRef ws As Worksheet, ByVal password$)
    ' ф-ция снимает пароль с листа
    ws.Unprotect Password:=password
End Function