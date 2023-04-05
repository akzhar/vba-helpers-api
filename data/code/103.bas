Attribute VB_Name = "Helper103"
Option Explicit

' Function GetAdGroupMembers(ByVal groupName$) As String
    
'     Dim members: members = GetADInfo("groups", "samAccountName", groupName, "member")
'     Dim arr(), userLogin$
    
'     Dim i&
'     For i = LBound(members) To UBound(members)
    
'         userLogin = GetAdUserLogin(members(i))
'         Call AddToArr(arr, userLogin) ' @dependency: 1.bas
    
'     Next i
    
'     GetAdGroupMembers = Join(arr, "," & Chr(10))

' End Function

' Function GetAdUserLogin(ByVal userCn$) As String

'     GetAdUserLogin = GetADInfo("users", "cn", userCn, "samAccountName")

' End Function

Const LDAP_DOMAIN$ = "LDAP://dc=sub, dc=example, dc=com"

Function GetADInfo(ByVal objectClass$, ByVal searchByAttr$, ByVal searchString$, ByVal returnAttr$) As Variant()
    ' Gets Active Directory attribute value (array of matches)
    
    Dim objConnection As Object: Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"
    Dim adoCommand As Object: Set adoCommand = CreateObject("ADODB.Command")
    adoCommand.ActiveConnection = objConnection
    
    Dim ldapFilter$
    
    Select Case objectClass
        Case "users"
            ldapFilter = "(&(|(objectClass=user)(objectClass=person))(!(objectClass=computer))(!(objectClass=group)))"
        Case "groups"
            ldapFilter = "(&(objectClass=group)(!(objectClass=computer))(!(objectClass=user))(!(objectClass=person)))"
        Case "computers"
            ldapFilter = "(objectClass=computer)"
    End Select
        
    adoCommand.CommandText = "<" & LDAP_DOMAIN & ">;" _
    & "(&" & ldapFilter & "(" & searchByAttr & "=" & searchString & "));" _
    & searchByAttr & "," & returnAttr & ";subtree"
    
    Dim objRecordSet: Set objRecordSet = adoCommand.Execute
    
    Dim arr()
    
    If objRecordSet.RecordCount > 0 Then
        Dim i&, item$
        For i = 1 To objRecordSet.RecordCount
            item = objRecordSet.Fields(returnAttr)
            Call AddToArr(arr, item) ' @dependency: 1.bas
            objRecordSet.MoveNext
        Next i
    End If
    
    GetADInfo = arr
    
    objConnection.Close
    Set objRecordSet = Nothing
    Set adoCommand = Nothing
    Set objConnection = Nothing

End Function

Function GetADInfo(ByVal objectClass$, ByVal searchByAttr$, ByVal searchString$, ByVal returnAttr$) As Variant
    ' Gets Active Directory attribute value (first match only)
    
    Dim objConnection As Object: Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Open "Provider=ADsDSOObject;"
    Dim adoCommand As Object: Set adoCommand = CreateObject("ADODB.Command")
    adoCommand.ActiveConnection = objConnection
    
    Dim ldapFilter$
    
    Select Case objectClass
        Case "users"
            ldapFilter = "(&(|(objectClass=user)(objectClass=person))(!(objectClass=computer))(!(objectClass=group)))"
        Case "groups"
            ldapFilter = "(&(objectClass=group)(!(objectClass=computer))(!(objectClass=user))(!(objectClass=person)))"
        Case "computers"
            ldapFilter = "(objectClass=computer)"
    End Select
        
    adoCommand.CommandText = "<" & LDAP_DOMAIN & ">;" _
    & "(&" & ldapFilter & "(" & searchByAttr & "=" & searchString & "));" _
    & searchByAttr & "," & returnAttr & ";subtree"
    
    Dim objRecordSet: Set objRecordSet = adoCommand.Execute
    
    If objRecordSet.RecordCount = 0 Then
        GetADInfo = "not found"
    Else
        GetADInfo = objRecordSet.Fields(returnAttr)
    End If
    
    objConnection.Close
    Set objRecordSet = Nothing
    Set adoCommand = Nothing
    Set objConnection = Nothing

End Function