Attribute VB_Name = "Helper21"
Option Explicit

Function UnixTime2Date(ByVal unixDate$) As Date
    ' Converts Unix 13-digit timestamp to date
    Dim sec&: sec = Val(unixDate / 1000)
    Dim dStart As Date: dStart = #1/1/1970#
    UnixTime2Date = DateAdd("s", sec, dStart)
End Function