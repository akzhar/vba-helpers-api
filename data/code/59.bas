Attribute VB_Name = "Helper59"
Option Explicit

Function ProtectWs(ByVal flag As Boolean, ByRef ws As Worksheet, ByVal password$)
    ' ф-ция ставит / снимает пароль с листа
    If flag Then
        With ws
            .Protect _
                Password:=password, _
                AllowFiltering:=True ' автофильтр вкл
            .EnableSelection = xlNoRestrictions
        End With
    Else
        ws.Unprotect Password:=password
    End if
End Function