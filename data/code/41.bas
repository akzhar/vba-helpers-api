Attribute VB_Name = "Helper41"
Option Explicit

Function ShowInterface(ByVal flag As Boolean)
    ' ф-ция скрывает / показывает пользовательский интерфейс Excel
    With Application
        .ScreenUpdating = False
        .Caption = IIf(flag = True, Empty, "")
        .DisplayStatusBar = flag: .DisplayFormulaBar = flag
        Dim iCommandBar As CommandBar
        For Each iCommandBar In .CommandBars
            iCommandBar.Enabled = flag
        Next iCommandBar
        With .ActiveWindow
            .Caption = IIf(flag = True, .Parent.Name, "")
            .DisplayHeadings = flag: .DisplayGridlines = flag
            '.DisplayHorizontalScrollBar = flag: .DisplayVerticalScrollBar = flag
            '.DisplayWorkbookTabs = flag
        End With
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", " & flag & ")"
        .ScreenUpdating = True
    End With
End Function