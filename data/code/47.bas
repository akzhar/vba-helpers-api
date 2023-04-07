Attribute VB_Name = "VbaHelper_CreateEvent"
Option Explicit

Function CreateEvent(ByVal subject$, ByVal body$, ByVal startDate As Date, Optional ByVal recurrenceType$ = "")
    ' Creates an event current user's Outlook calendar

    If subject = "" Then
        MsgBox "Error: empty subject", vbExclamation
        Exit Function
    End If

    If startDate < Date Then
        MsgBox "Error: date < today", vbExclamation
        Exit Function
    End If
    
    Const CALENDAR_FOLDER_TYPE& = 9 ' 9 = main calendar
    Const APPOINTMENT_TYPE& = 1 ' 1 = appointment

    Dim outlookObj As Object: Set outlookObj = CreateObject("Outlook.Application")
    Dim namespaceObj As Object: Set namespaceObj = outlookObj.GetNamespace("MAPI")
    Dim calendarItemsObj As Object: Set calendarItemsObj = namespaceObj.GetDefaultFolder(CALENDAR_FOLDER_TYPE).Items
    Dim calendarObj As Object: Set calendarObj = namespaceObj.GetDefaultFolder(CALENDAR_FOLDER_TYPE)
    
    Dim calendarItem As Object
    
    For Each calendarItem In calendarItemsObj
        If calendarItem.subject = subject Then
            MsgBox "Error: an event with this header already exists", vbExclamation
            Exit Function
        End If
    Next calendarItem

    Dim newEventItem As Object: Set newEventItem = calendarObj.Items.Add(APPOINTMENT_TYPE)
    
    newEventItem.subject = subject
    newEventItem.body = IIf(body <> "", body & vbLf & vbLf, "") & "Created by macros"
    newEventItem.Start = startDate
    newEventItem.AllDayEvent = True
    
    If recurrenceType <> "" Then
                
        Dim recurrencePattern As Object: Set recurrencePattern = newAppointmentItem.GetRecurrencePattern
        recurrencePattern.PatternStartDate = startDate
        
        Select Case recurrenceType
            Case "Daily"
                recurrencePattern.recurrenceType = 0
            Case "Weekly"
                recurrencePattern.recurrenceType = 1
            Case "Monthly"
                recurrencePattern.recurrenceType = 3
            Case "Annual"
                recurrencePattern.recurrenceType = 5
        End Select
        
    End If
    
    newEventItem.Save
    
    Set outlookObj = Nothing
    Set namespaceObj = Nothing
    Set calendarItemsObj = Nothing
    Set calendarObj = Nothing
    Set calendarItem = Nothing
    Set newEventItem = Nothing
    Set recurrencePattern = Nothing
    
    MsgBox "Event has been created: " & subject & IIf(recurrenceType <> "", " (reccurence " & recurrenceType & ")", ""), vbInformation
    
End Function