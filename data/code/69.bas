Attribute VB_Name = "Helper69"
Option Explicit

Function Includes(ByVal str$, ByVal subStr$) As Boolean
    ' ф-ция проверят вхождение подстроки в строку
    Includes = CBool(InStr(1, str, subStr, vbTextCompare) <> 0)
End Function