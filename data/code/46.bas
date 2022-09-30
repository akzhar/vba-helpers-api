Attribute VB_Name = "Helper46"
Option Explicit

Function SendEmail(ByVal subject$, ByVal body$, ByVal sendTo$, Optional ByVal copyTo$, Optional ByVal attachmentPath$ = "", Optional ByVal method$ = "Show", Optional ByVal importance$ = "Low")
    ' Sends email in Outlook behalf of the current user

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

    Dim messageEnding$
        
    With Mail

        If Not IsNull(copyTo) Then
            .CC = copyTo
        End If
        .To = sendTo
        .importance = importanceType
        .subject = subject
        .body = body

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
        
    Set Mail = Nothing
    Set Outlook = Nothing

    MsgBox "Email has been " & messageEnding, vbInformation

End Function