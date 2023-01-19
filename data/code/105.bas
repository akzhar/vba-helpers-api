Attribute VB_Name = "Helper105"
Option Explicit

Sub ToggleExpandCollapse()
    ' Expand / collapse rows and columns on active sheet
    
    Static isExpanded As Boolean

    ThisWorkbook.ActiveSheet.Outline.ShowLevels _
        Rowlevels:=IIf(isExpanded, 1, 8), _
        ColumnLevels:=IIf(isExpanded, 1, 8)
    
    isExpanded = Not isExpanded

End Sub
