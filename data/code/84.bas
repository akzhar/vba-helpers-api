Attribute VB_Name = "Helper84"
Option Explicit

Public Const VAR_DELIMITER$ = "%"

Private Function GetRegExpPattern(ByVal startRow$, ByVal endRow$) As String
    ' Returns regexp pattern
    GetRegExpPattern = "(?:" & startRow & "$\s)([\S\s]+?)(?:" & endRow & "$\s?)"
End Function

Private Function GetStubsArr(ByVal template$) As Variant()
    ' Returns array of stubs from the template
    Dim stubsArr(): stubsArr = GetRegExpMatches(template, VAR_DELIMITER & "\w*" & VAR_DELIMITER) ' @(id 60)
    stubsArr = GetUniqueArr(stubsArr) ' @(id 10)
    GetStubsArr = stubsArr
End Function

Private Function GetBlockContent(ByVal template$, ByVal startRow$, ByVal endRow$) As String
    ' Returns block between startRow и endRow
    Dim matches(): matches = GetRegExpMatches(template, GetRegExpPattern(startRow, endRow)) ' @(id 60)
    GetBlockContent = ""
    If GetArrLength(matches) > 0 Then ' @(id 2)
        GetBlockContent = SliceString(template, InStr(1, template, startRow) + Len(startRow), InStr(1, template, endRow) - 1) ' @(id 72)
    End If
End Function

Private Sub InsertNewRow(ByRef template$, ByRef rowToInsert$, ByVal afterRow$)
    ' Insert new row after ifStartRow
    rowToInsert = IIf(Right(rowToInsert, 1) = vbLf, Left(rowToInsert, Len(rowToInsert) - 1), rowToInsert)
    rowToInsert = IIf(Left(rowToInsert, 1) = vbLf, Right(rowToInsert, Len(rowToInsert) - 1), rowToInsert)
    Dim afterRowPos&: afterRowPos = InStr(1, template, afterRow) + Len(afterRow)
    template = Left(template, afterRowPos - 1) & vbLf & rowToInsert & Mid(template, afterRowPos)
End Sub

Private Function GetIfStartRow(ByVal dataKey$) As String
    GetIfStartRow = "<!-- IF HAS " & VAR_DELIMITER & dataKey & VAR_DELIMITER & " -->"
End Function

Private Function GetIfEndRow(ByVal dataKey$) As String
    GetIfEndRow = "<!-- END IF " & VAR_DELIMITER & dataKey & VAR_DELIMITER & " -->"
End Function

Private Function GetLoopStartRow(ByVal dataKey$) As String
    GetLoopStartRow = "<!-- LOOP EACH " & VAR_DELIMITER & dataKey & VAR_DELIMITER & " -->"
End Function

Private Function GetLoopEndRow(ByVal dataKey$) As String
    GetLoopEndRow = "<!-- STOP LOOP " & VAR_DELIMITER & dataKey & VAR_DELIMITER & " -->"
End Function

Private Function FillTemplateWithData(ByVal template$, ByRef stubsArr(), ByRef dataMap As Scripting.Dictionary) As String
    ' Fills in the template with data
    
    Dim stubIndex&
    For stubIndex = LBound(stubsArr) To UBound(stubsArr)
        
        Dim subStubsArr()
        Dim dataObj As Scripting.Dictionary
        Dim blockStartRow$, blockEndRow$, rowToInsert$, templateRow$, blockContent$
        Dim isCondition As Boolean, isLoop As Boolean, isValidData As Boolean
        
        Dim dataKey$: dataKey = stubsArr(stubIndex)
        dataKey = Replace(dataKey, VAR_DELIMITER, "")
        
        blockStartRow = GetIfStartRow(dataKey)
        blockEndRow = GetIfEndRow(dataKey)
        blockContent = GetBlockContent(template, blockStartRow, blockEndRow)
        
        isCondition = CBool(blockContent <> "" And InStr(1, blockContent, VAR_DELIMITER) > 0)
        isValidData = CBool(TypeName(dataMap(dataKey)) = "Dictionary")
            
        If isCondition Then
            
            If Not IsEmpty(dataMap(dataKey)) And isValidData Then
            
                Set dataObj = dataMap(dataKey)
                                
                subStubsArr = GetStubsArr(blockContent)
                rowToInsert = FillTemplateWithData(blockContent, subStubsArr, dataObj)
                Call InsertNewRow(template, rowToInsert, blockStartRow)
            
            End If
                    
            ' Delete IF template block
            template = Replace(template, blockContent, "")
                
        End If
        
        blockStartRow = GetLoopStartRow(dataKey)
        blockEndRow = GetLoopEndRow(dataKey)
        blockContent = GetBlockContent(template, blockStartRow, blockEndRow)
        
        isLoop = CBool(blockContent <> "" And InStr(1, blockContent, VAR_DELIMITER) > 0)
        isValidData = IsArray(dataMap(dataKey))
        
        If isLoop Then
            
            If Not IsEmpty(dataMap(dataKey)) And isValidData Then
            
                Dim arrElem() As Scripting.Dictionary: arrElem = dataMap(dataKey)
                       
                Dim arrIndex&
                For arrIndex = UBound(arrElem) To 0 Step -1
                    
                    Set dataObj = arrElem(arrIndex)
                    
                    If Not dataObj Is Nothing Then
                        subStubsArr = GetStubsArr(blockContent)
                        rowToInsert = FillTemplateWithData(blockContent, subStubsArr, dataObj)
                        Call InsertNewRow(template, rowToInsert, blockStartRow)
                    End If
                    
                Next arrIndex
                
            End If
            
            ' Delete LOOP template block
            template = Replace(template, blockContent, "")
            
        End If
        
        If Not isCondition And Not isLoop Then

            template = Replace(template, VAR_DELIMITER & dataKey & VAR_DELIMITER, dataMap(dataKey))
            
        End If
        
    Next stubIndex
        
    FillTemplateWithData = template
    
End Function