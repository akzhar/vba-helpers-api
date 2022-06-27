Attribute VB_Name = "Helper70"
Option Explicit

Function GetFirstWordAfter(ByVal searchWord$, ByVal str$) As String
    ' ф-ция возвращает первое слово в строке после искомого слова
    
    Dim wordAfter$
    
    wordAfter = Mid(str, InStr(1, str, " " & searchWord & " ", vbTextCompare) + Len(searchWord))
    wordAfter = Mid(wordAfter, InStr(1, wordAfter, " ", vbTextCompare))
    wordAfter = Trim(wordAfter)
    
    If Includes(wordAfter, " ") Then
        wordAfter = Mid(wordAfter, 1, InStr(1, wordAfter, " ", vbTextCompare) - 1)
    End If
    
    GetFirstWordAfter = wordAfter

End Function