Attribute VB_Name = "VbaHelper_ExpandCollapseRowCol"
Option Explicit

Function ExpandCollapseRowCol(ByRef ws As Worksheet, Optional ByVal mode$)
    ' Expandes / collapses grouped rows and columns on the sheet

    Static isExpanded As Boolean
    
    Dim isToggleMode As Boolean: isToggleMode = CBool(mode = "")

    ws.Outline.ShowLevels _
        Rowlevels:=IIf(LCase(mode) = "collapse", 1, IIf(LCase(mode) = "expand", 8, IIf((isToggleMode And isExpanded), 1, 8))), _
        ColumnLevels:=IIf(LCase(mode) = "collapse", 1, IIf(LCase(mode) = "expand", 8, IIf((isToggleMode And isExpanded), 1, 8)))
    
    isExpanded = Not isExpanded

End Function