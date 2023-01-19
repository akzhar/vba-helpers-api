Attribute VB_Name = "Helper46"
Option Explicit

Sub CreateEmail( _
    ByVal subject$, _
    ByVal message$, _
    ByVal sendTo$, _
    Optional ByVal copyTo$, _
    Optional ByVal hidenCopyTo$, _
    Optional ByVal replyTo$, _
    Optional ByVal isHtml As Boolean = False, _
    Optional ByVal attachmentPath$ = "", _
    Optional ByVal method$ = "Show", _
    Optional ByVal importance$ = "Medium" _
)
    ' Creates email in Outlook behalf of the current user

    Const OUTLOOK_ITEM_TYPE& = 0
    
    Dim importanceType&
    
    Select Case importance
        Case "High"
            importanceType = 2
        Case "Medium"
            importanceType = 1
        Case "Low"
            importanceType = 0
        Case Else
            importanceType = 0
    End Select
        
    Dim Outlook As Object: Set Outlook = CreateObject("Outlook.Application")
    Dim Mail As Object: Set Mail = Outlook.CreateItem(OUTLOOK_ITEM_TYPE)
    Dim Inspector As Object: Set Inspector = Mail.GetInspector
        
    Dim messageEnding$
        
    With Mail

        If Not IsNull(copyTo) And copyTo <> "" Then .CC = copyTo
        If Not IsNull(hidenCopyTo) And hidenCopyTo <> "" Then .Bcc = hidenCopyTo
        If Not IsNull(replyTo) And replyTo <> "" Then
            .ReplyRecipients.Add replyTo
            .ReplyRecipients.Add Utils.GetCurrentUserEmail() '(@id 45)
            If Not IsNull(copyTo) And copyTo <> "" Then .ReplyRecipients.Add copyTo
        End If
        .To = sendTo
        .importance = importanceType
        .subject = subject
        If isHtml Then
            .HTMLbody = message
        Else
            .body = message
        End If
    
        If attachmentPath <> "" Then
            .Attachments.Add (attachmentPath)
        End If
          
        Select Case method
            Case "Show"
                .Display
                messageEnding = "created"
            Case "Save"
                .Save
                messageEnding = "saved to Drafts folder"
            Case "Send"
                .Send
                messageEnding = "sent"
            Case Else
                .Display
                messageEnding = "created"
        End Select
    
    End With
        
    Set Mail = Nothing: Set Outlook = Nothing

    'Debug.Print "Email has been " & messageEnding

End Sub