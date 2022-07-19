Attribute VB_Name = "Helper47"
Option Explicit

Function CreateAppointment(ByVal subject$, ByVal body$, ByVal startDate As Date, Optional ByVal recurrenceType$ = "")
    ' ф-ция создает в календаре текущего пользователя новое событие

    If subject = "" Then
        MsgBox "Ошибка: пустой заголовок", vbExclamation
        Exit Function
    End If

    ' события в календаре не создаются для дат из прошлого
    If startDate < Date Then
        MsgBox "Ошибка: дата < сегодня", vbExclamation
        Exit Function
    End If
    
    Const CALENDAR_FOLDER_TYPE& = 9 ' 9 = main calendar
    Const APPOINTMENT_TYPE& = 1 ' 1 = appointment

    Dim outlookObj As Object: Set outlookObj = CreateObject("Outlook.Application")
    Dim namespaceObj As Object: Set namespaceObj = outlookObj.GetNamespace("MAPI")
    Dim appointments As Object: Set appointments = namespaceObj.GetDefaultFolder(CALENDAR_FOLDER_TYPE).Items
    Dim calendarObj As Object: Set calendarObj = namespaceObj.GetDefaultFolder(CALENDAR_FOLDER_TYPE)
        
    ' проверка наличия повторов в календаре
    Dim appointmentItem As Object
    
    For Each appointmentItem In appointments
        
        If appointmentItem.subject = subject Then
            MsgBox "Ошибка: событие с таким заголовком уже существует", vbExclamation
            Exit Function
        End If
    
    Next appointmentItem

    Dim newAppointmentItem As Object: Set newAppointmentItem = calendarObj.Items.Add(APPOINTMENT_TYPE)
    
    newAppointmentItem.subject = subject
    newAppointmentItem.body = IIf(body <> "", body & vbLf & vbLf, "") & "Создано с помощью макроса"
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
    
    MsgBox "Событие в календаре создано: " & subject & IIf(recurrenceType <> "", " (повтор " & recurrenceType & ")", ""), vbInformation
    
End Function