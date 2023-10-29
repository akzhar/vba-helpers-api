Attribute VB_Name = "VbaHelper_ReadClosedExcelFile"
Option Explicit

Function ReadClosedExcelFile(ByVal wbPath$, ByVal wbName$, ByVal wsName$, ByVal cellAddress$) As Variant
    ' Reads single cell value from Excel file without openning it
    
    If Right(wbPath, 1) <> "\" Then
        wbPath = wbPath & "\"
    End If
    
    ReadClosedExcelFile = CStr(ExecuteExcel4Macro("'" & wbPath & "[" & wbName & "]" & wsName & "'!" & Range(cellAddress).Address(ReferenceStyle:=xlR1C1)))

End Function