Attribute VB_Name = "VbaHelper_AdUacDecoder"
Option Explicit

' 31 bit properties (index 0 to 30)
Private Const UAC_PROPERTIES As String = "SCRIPT,ACCOUNTDISABLE,RESERVED,HOMEDIR_REQUIRED,LOCKOUT,PASSWD_NOTREQD,PASSWD_CANT_CHANGE,ENCRYPTED_TEXT_PWD_ALLOWED,TEMP_DUPLICATE_ACCOUNT,NORMAL_ACCOUNT,RESERVED,INTERDOMAIN_TRUST_ACCOUNT,WORKSTATION_TRUST_ACCOUNT,SERVER_TRUST_ACCOUNT,RESERVED,RESERVED,DONT_EXPIRE_PASSWORD,MNS_LOGON_ACCOUNT,SMARTCARD_REQUIRED,TRUSTED_FOR_DELEGATION,NOT_DELEGATED,USE_DES_KEY_ONLY,DONT_REQ_PREAUTH,PASSWORD_EXPIRED,TRUSTED_TO_AUTH_FOR_DELEGATION,RESERVED,PARTIAL_SECRETS_ACCOUNT,RESERVED,RESERVED,RESERVED,RESERVED"

' These values will be highlighted separately
Private Const ALERT_VALUES As String = "LOCKOUT,PASSWORD_EXPIRED,ACCOUNTDISABLE,DONT_EXPIRE_PASSWORD"

' Function accepts a decimal number
' Function returns a binary number - 31 digits (LSB first)
Private Function ToBase2(base10 As Long) As String

    Dim value&, result$, i As Integer
    
    value = base10
    result = ""
    
    For i = 1 To 31
        result = (value Mod 2) & result
        value = value \ 2
    Next i
    
    ToBase2 = result

End Function

' Function accepts userAccountControl value (decimal number)
' Function returns its human-readable value
Function AdUacDecoder(ByVal UACvalue As Long) As String

    Dim binaryString$, result$, prop$, i As Integer
    Dim propArray() As String, alertArray() As String
    
    ' Initialize arrays
    propArray = Split(UAC_PROPERTIES, ",")
    alertArray = Split(ALERT_VALUES, ",")
    
    ' Get binary representation (31 digits)
    binaryString = ToBase2(UACvalue)
    
    ' Process each bit from right to left (LSB first)
    For i = 31 To 1 Step -1
        If Mid(binaryString, i, 1) = "1" Then
            prop = propArray(31 - i)  ' Correct index mapping
            
            ' Check if this property should be highlighted
            If Not IsError(Application.Match(prop, alertArray, 0)) Then
                prop = "!!! " & prop
            End If
            
            result = result & prop & "; "
        End If
    Next i
    
    ' Remove trailing semicolon and space if needed
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 2)
    End If
    
    AdUacDecoder = result

End Function