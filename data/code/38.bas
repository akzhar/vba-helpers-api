Attribute VB_Name = "Helper38"
Option Explicit

Function Hex2Rgb(ByVal hexColor$) As Long
    ' ф-ция конвертирует HEX в RGB long value
    Dim R$, G$, B$
    hexColor = Replace(hexColor, "#", "")
    R = Val("&H" & Mid(hexColor, 1, 2))
    G = Val("&H" & Mid(hexColor, 3, 2))
    B = Val("&H" & Mid(hexColor, 5, 2))
    Hex2Rgb = RGB(R, G, B)
End Function
