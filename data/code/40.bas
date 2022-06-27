Attribute VB_Name = "Helper40"
Option Explicit

Const PROXY_URL$ = "xxx.proxy.su"
Const API_URL$ = "https://websbor.gks.ru/webstat/api/gs/organizations"
Dim HttpCodes As Object

Function init()
    Set HttpCodes = CreateObject("Scripting.Dictionary")
    HttpCodes("200") = "OK"
    HttpCodes("400") = "BAD REQUEST"
    HttpCodes("404") = "NOT FOUND"
    HttpCodes("500") = "SERVER ERROR"
End Function

Function HtppQuery(ByVal method$, ByVal url$, ByVal contentType$, Optional ByVal reqBody$) As Variant
    ' ф-ция выполняет HTTP запрос
    
    Dim req As Object: Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
          
    With req
        .Open method, url, False
        If PROXY_ON Then
            .setProxy 2, PROXY_URL, ""
        End If
        .setRequestHeader "Content-Type", contentType & "; charset=UTF-8"
        .send reqBody
    End With
    
    If req.Status <> "200" Then
        MsgBox "Запрос завершился неудачно." _
        & vbLf & vbLf & "Статус ответа: " & req.Status & HTTP_CODES(req.Status), vbExclamation
        Exit Function
    End If
    
    If contentType = "application/json" Then
        Set Query = JsonConverter.ParseJson(req.responseText) ' @(id 42)
    Else
        Set Query = req.responseText
    End If
            
End Function

' Sub test()
'     Dim json As Object, jsonItem As Object, inn$

'     inn = "7736207543"
        
'     Set json = Http.Query("POST", API_URL, "application/json", "{Inn:" & """" & inn & """" & "}")

'     If json Is Nothing Or json.Count = 0 Then
'         Debug.Print("There is nothing found related to INN " & inn)
'     End If

'     For Each jsonItem In json
'         Debug.Print("Стат форма " & jsonItem("index"))
'         Debug.Print("Срок подачи " & jsonItem("end_time"))
'     Next jsonItem
' End Sub