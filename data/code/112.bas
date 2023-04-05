Attribute VB_Name = "Helper112"
Option Explicit

Function UnzipFile(ByVal zippedFilePath$) As String
    ' Extract all files from a zip archive into the Temp folder
    
    Dim ts$: ts = Utils.GetTimeStamp() ' @dependency: 99.bas
    ts = Replace(ts, "-", "_")
    ts = Replace(ts, ":", "_")
    Dim folderName$: folderName = "unziped_" & ts
    
    Call Utils.CreateFolder(Environ("temp"), folderName) ' @dependency: 22.bas
    
    Dim pathToUnzip$: pathToUnzip = Environ("temp") & Application.PathSeparator & folderName
    
    Dim ShellApp As Object: Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace((pathToUnzip)).CopyHere ShellApp.Namespace((zippedFilePath)).Items
    
    UnzipFile = pathToUnzip
    
End Function