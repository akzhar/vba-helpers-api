Attribute VB_Name = "Helper72"
Option Explicit

Function SliceString(ByVal textString$, ByVal beginIndex&, Optional ByVal endIndex&) As String
    ' ф-ция извлекает подстроку из строки начиная с beginIndex и заканчивая endIndex (нумерация начинается с 1)
    
    If Len(textString) = 0 Then
        SliceString = textString
        Exit Function
    End If

    beginIndex = IIf(beginIndex = 0, 1, beginIndex)
    beginIndex = IIf(beginIndex < 0, Len(textString) - beginIndex, beginIndex)
    endIndex = IIf(IsNull(endIndex) Or endIndex = 0 Or endIndex > Len(textString), Len(textString), endIndex)
    endIndex = IIf(endIndex < 1, Len(textString) + endIndex, endIndex)
    
    SliceString = Mid(textString, beginIndex, endIndex - beginIndex + 1)
End Function