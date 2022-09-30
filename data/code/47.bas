Attribute VB_Name = "Helper47"
Option Explicit

Function CreateAppointment(ByVal subject$, ByVal body$, ByVal startDate As Date, Optional ByVal recurrenceType$ = "")
    ' Creates an appointment current user's Outlook calendar

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
    Dim appointments As Object: Set appointments = namespaceObj.GetDefaultFolder(CALENDAR_FOLDER_TYPE).Items
    Dim calendarObj As Object: Set calendarObj = namespaceObj.GetDefaultFolder(CALENDAR_FOLDER_TYPE)
    
    Dim appointmentItem As Object
    
    For Each appointmentItem In appointments
        If appointmentItem.subject = subject Then
            MsgBox "Error: an event with this header already exists", vbExclamation
            Exit Function
        End If
    Next appointmentItem

    Dim newAppointmentItem As Object: Set newAppointmentItem = calendarObj.Items.Add(APPOINTMENT_TYPE)
    
    newAppointmentItem.subject = subject
    newAppointmentItem.body = IIf(body <> "", body & vbLf & vbLf, "") & "Created by macros"
    newAppointmentItem.Start = startDate
    newAppointmentItem.AllDayEvent = True
    
    ' периодический повтор задачи
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
    
    newAppointmentItem.Save
    
    Set outlookObj = Nothing
    Set namespaceObj = Nothing
    Set appointments = Nothing
    Set calendarObj = Nothing
    Set appointmentItem = Nothing
    Set newAppointmentItem = Nothing
    Set recurrencePattern = Nothing
    
    MsgBox "Appointment has been created: " & subject & IIf(recurrenceType <> "", " (reccurence " & recurrenceType & ")", ""), vbInformation
    
End Function