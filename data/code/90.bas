Attribute VB_Name = "Helper90"
Option Explicit

Function SplitToChunks(ByVal text$, ByVal numOfChars&) As String()
    ' ф-ция разбивает строку на массив строк заданной длины
    Dim chunks() As String
    Dim chunk$, count&: count = 0
    Do While Len(text)
        ReDim Preserve chunks(count)
        chunk = Left(text, numOfChars)
        chunks(count) = chunk
        text = Mid(text, Len(chunk) + 1)
        count = count + 1
    Loop
    SplitToChunks = chunks
End Function