Attribute VB_Name = "VbaHelper_LogInfo"
Option Explicit

Function LogInfo(ByVal logMessage$, Optional ByVal logFileName$ = "logs")
    ' Writes new line in log file with current timestamp

    Dim logFilePath$: logFilePath = ThisWorkbook.Path & Application.PathSeparator & logFileName & ".log"
    Dim logFileNum As Integer: logFileNum = FreeFile
    
    ' If there is no log file create it
    Open logFilePath For Append As logFileNum
    
    ' Write new line in the end
    Print #logFileNum, GetTimeStamp() & " - " & logMessage ' @dependency: 99.bas
    
    ' Close log file
    Close logFileNum
End Function