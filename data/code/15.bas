Attribute VB_Name = "Helper15"
Option Explicit

Function GetDateSeparator()
    ' Gets date separator based on user region settings (locale)
    GetDateSeparator = Excel.Application.International(XlApplicationInternational.xlDateSeparator)
End Function