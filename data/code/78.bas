Attribute VB_Name = "VbaHelper_ValidateInput"
Option Explicit

Function ValidateInput(ByVal textRequest$, ByVal checkPattern$, ByVal textPattern$, Optional ByVal textWarning$, Optional ByVal defaultInput$) As String
    ' Prompts you to enter a value and checks its correctness using specified pattern

    Dim inputData$
    
    If textWarning <> "" Then textWarning = vbLf & vbLf & "Warning!" & vbLf & textWarning

SelectData:

    inputData = Trim(InputBox(textRequest & vbLf & vbLf & "Follow the format: " & textPattern & textWarning, "Please type the value", defaultInput))

    If inputData = "" Then
        If MsgBox("Empty value." & vbLf & vbLf & "Want to type again?", vbYesNo, "Error") = vbYes Then
            GoTo SelectData
        Else
            Exit Function
        End If
    End If

    If Not inputData Like checkPattern Then
        If MsgBox("Validation failed: " & textPattern & "." & vbLf & vbLf & "Want to type again?", vbYesNo, "Error") = vbYes Then
            GoTo SelectData
        Else
            Exit Function
        End If
    End If

    ValidateInput = inputData
End Function