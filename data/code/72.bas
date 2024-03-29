Attribute VB_Name = "VbaHelper_SliceString"
Option Explicit

Function SliceString(ByVal textString$, ByVal beginIndex&, Optional ByVal endIndex&) As String
    ' Extracts a substring from the string
    
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