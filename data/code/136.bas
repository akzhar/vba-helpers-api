Attribute VB_Name = "VbaHelper_CopyExcelFiles"
Option Explicit

Function CopyExcelFiles(ByVal sourceFolderPath$, ByVal targetFolderPath$) As Long
    ' Copies all Excel files from source folder to target folder
    ' Returns number of copied files

    Dim fileCopiedCounter&: fileCopiedCounter = 0

    If Right(sourceFolderPath, 1) <> "\" Then sourceFolderPath = sourceFolderPath & "\"
    If Right(targetFolderPath, 1) <> "\" Then targetFolderPath = targetFolderPath & "\"

    Dim fileName$: fileName = Dir(sourceFolderPath & "*.xls*")

    Do While fileName <> ""
        Call FileCopy(Source:=sourceFolderPath & fileName, Destination:=targetFolderPath & fileName)
        fileCopiedCounter = fileCopiedCounter + 1
        fileName = Dir
    Loop

    CopyExcelFiles = fileCopiedCounter
    
End Function