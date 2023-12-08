Attribute VB_Name = "VbaHelper_SplitToChunks"
Option Explicit

Function SplitToChunks(ByVal text$, ByVal chunkSize&) As String()
    ' Splits the specified string into an array of strings of a given length
    Dim chunks() As String
    Dim chunk$, count&: count = 0
    Do While Len(text)
        ReDim Preserve chunks(count)
        chunk = Left(text, chunkSize)
        chunks(count) = chunk
        text = Mid(text, Len(chunk) + 1)
        count = count + 1
    Loop
    SplitToChunks = chunks
End Function