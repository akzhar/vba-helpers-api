Attribute VB_Name = "VbaHelper_Rgb2Long"
Option Explicit

Function Rgb2Long(ByVal R&, ByVal G&, ByVal B&) As Long
    ' Converts RGB color to a Long Excel value
    Rgb2Long = (B * 65536) + (G * 256) + R
End Function