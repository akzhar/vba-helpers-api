Attribute VB_Name = "Helper44"
Option Explicit

Function LogInfo(ByVal logMessage$)
    ' ф-ция записывает сообщение в логфайл с указанием даты и времени
    
    Const LOG_FILE_NAME$ = "Log.txt"

    Dim timeStamp$: timeStamp = CStr(Format(Now, "dd.mm.yyyy hh:mm:ss"))
    Dim logFilePath$: logFilePath = ThisWorkbook.Path & Application.PathSeparator & LOG_FILE_NAME
    Dim logFileNum As Integer: logFileNum = FreeFile
    
    ' создает файл если его нет
    Open logFilePath For Append As logFileNum
    
    ' добавляет информацию в конец файла и закрывает его
    Print #logFileNum, timeStamp & " - " & logMessage
    
    Close logFileNum
End Function