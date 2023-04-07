Attribute VB_Name = "VbaHelper_Hex2Long"
Option Explicit

Function Hex2Long(ByVal hexColor$) As Long
    ' Converts HEX color to Long Excel value
    Dim R$, G$, B$
    hexColor = Replace(hexColor, "#", "")
    R = Val("&H" & Mid(hexColor, 1, 2))
    G = Val("&H" & Mid(hexColor, 3, 2))
    B = Val("&H" & Mid(hexColor, 5, 2))
    Hex2Long = RGB(R, G, B)
End Function
