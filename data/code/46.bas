Attribute VB_Name = "Helper46"
Option Explicit

'@importanceType - "Высокая", "Средняя", "Низкая"
'@method = "Показать", "Сохранить", "Отправить"
Function SendEmail(ByVal subject$, ByVal body$, ByVal sendTo$, Optional ByVal copyTo$, Optional ByVal attachmentPath$ = "", Optional ByVal method$ = "Показать перед отправкой", Optional ByVal importanceType$ = "Низкая")
    ' ф-ция отправляет письмо в Outlook

    Const OUTLOOK_ITEM_TYPE& = 0
    
    Dim importance&
    
    Select Case importanceType
        Case "Высокая"
            importance = 2
        Case "Средняя"
            importance = 1
        Case "Низкая"
            importance = 0
        Case Else
            importance = 0
    End Select
        
    Dim Outlook As Object: Set Outlook = CreateObject("Outlook.Application")
    Dim Mail As Object: Set Mail = Outlook.CreateItem(OUTLOOK_ITEM_TYPE)

    Dim messageEnding$
        
    With Mail

        If Not IsNull(copyTo) Then
            .CC = copyTo
        End If
        .To = sendTo
        .importance = importance
        .subject = subject
        .body = body

        If attachmentPath <> "" Then
            .Attachments.Add (attachmentPath)
        End If
          
        Select Case method
            Case "Показать"
                .Display
                messageEnding = "сформировано"
            Case "Сохранить"
                .Save
                messageEnding = "сохранено в папке Черновики / Drafts"
            Case "Отправить"
                .Send
                messageEnding = "отправлено"
            Case Else
                .Display
                messageEnding = "сформировано"
        End Select
    
    End With
        
    Set Mail = Nothing
    Set Outlook = Nothing

    MsgBox "Письмо " & messageEnding, vbInformation

End Function