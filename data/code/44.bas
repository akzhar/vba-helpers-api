Attribute VB_Name = "Helper44"
Option Explicit

Function LogInfo(ByVal logMessage$)
    ' Writes new line in log file with current timestamp
    
    Const LOG_FILE_NAME$ = "Log.txt"

    Dim timeStamp$: timeStamp = CStr(Format(Now, "dd.mm.yyyy hh:mm:ss"))
    Dim logFilePath$: logFilePath = ThisWorkbook.Path & Application.PathSeparator & LOG_FILE_NAME
    Dim logFileNum As Integer: logFileNum = FreeFile
    
    ' if there is no log file create it
    Open logFilePath For Append As logFileNum
    
    ' write new line in the end
    Print #logFileNum, timeStamp & " - " & logMessage
    
    ' close log file
    Close logFileNum
End Function