Attribute VB_Name = "Helper15"
Option Explicit

Function GetDateSeparator()
    ' ф-ция возвращает текущий разделитель дат в зависимости от региональных настроек пользователя
    GetDateSeparator = Excel.Application.International(XlApplicationInternational.xlDateSeparator)
End Function