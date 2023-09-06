Attribute VB_Name = "VbaHelper_ExpandCollapseRowCol"
Option Explicit

Function ExpandCollapseRowCol(ByRef ws As Worksheet, Optional ByVal flag = False)
    ' Expand / collapse rows and columns on the sheet
    ws.Outline.ShowLevels _
        Rowlevels:=IIf(flag, 8, 1), _
        ColumnLevels:=IIf(flag, 8, 1)
End Function