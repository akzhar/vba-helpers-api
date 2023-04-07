Attribute VB_Name = "VbaHelper_GetDateSeparator"
Option Explicit

Function GetDateSeparator()
    ' Gets date separator based on user region settings (locale)
    GetDateSeparator = Excel.Application.International(XlApplicationInternational.xlDateSeparator)
End Function