Attribute VB_Name = "VbaHelper_GetFirstWordAfter"
Option Explicit

Function GetFirstWordAfter(ByVal searchWord$, ByVal str$) As String
    ' Gets the 1st word from the text after the specified word
    
    Dim wordAfter$
    
    wordAfter = Mid(str, InStr(1, str, " " & searchWord & " ", vbTextCompare) + Len(searchWord))
    wordAfter = Mid(wordAfter, InStr(1, wordAfter, " ", vbTextCompare))
    wordAfter = Trim(wordAfter)
    
    If HasSubstring(wordAfter, " ") Then ' @dependency: 69.bas
        wordAfter = Mid(wordAfter, 1, InStr(1, wordAfter, " ", vbTextCompare) - 1)
    End If
    
    GetFirstWordAfter = wordAfter
End Function