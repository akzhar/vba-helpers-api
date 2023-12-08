Attribute VB_Name = "VbaHelper_GetDateFormatted"
Option Explicit

Function GetDateFormatted(ByVal d As Date, Optional dateFormat$ = "DD MMM YYYY", Optional lang$ = "eng") As String
    ' Returns string representation of specified date
    ' See codes for other languages here
    ' https://excel.tips.net/T003299_Specifying_a_Language_for_the_TEXT_Function.html
    ' https://www.myonlinetraininghub.com/excel-dates-displayed-in-different-languages
    Dim langCode$
    Select Case lang
      Case "eng":
        langCode = "eng-us" ' 0409
      Case "rus":
        langCode = "ru-ru" ' 0419	
      Case Else:
        Exit Function
    End Select

    GetDateFormatted = WorksheetFunction.Text(d, "[$-" & langCode & "]" & dateFormat)
End Function