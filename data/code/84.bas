Attribute VB_Name = "VbaHelper_InsertData2XmlTemplate"
Option Explicit

Private Const STUB_DELIMITER$ = "%"

Private Function GetIfStartRow(ByVal datakey$) As String
    GetIfStartRow = "<!-- IF HAS " & STUB_DELIMITER & datakey & STUB_DELIMITER & " -->"
End Function

Private Function GetIfEndRow(ByVal datakey$) As String
    GetIfEndRow = "<!-- END IF " & STUB_DELIMITER & datakey & STUB_DELIMITER & " -->"
End Function

Private Function GetLoopStartRow(ByVal datakey$) As String
    GetLoopStartRow = "<!-- LOOP EACH " & STUB_DELIMITER & datakey & STUB_DELIMITER & " -->"
End Function

Private Function GetLoopEndRow(ByVal datakey$) As String
    GetLoopEndRow = "<!-- STOP LOOP " & STUB_DELIMITER & datakey & STUB_DELIMITER & " -->"
End Function

Private Function RemoveStubDelimiters(ByVal datakey$) As String
    RemoveStubDelimiters = Replace(datakey, STUB_DELIMITER, "")
End Function

Private Function GetBlockContents(ByVal template$, ByVal startRowPattern$, ByVal endRowPattern$, Optional isGreedy As Boolean = False) As Variant()
    ' Returns 2D arr(N)(3) - N matches, 3 items for every match (startRow, block between, endRow)
    Dim regExpPattern$: regExpPattern = "(" & startRowPattern & "\s?)([\s\S]+" & IIf(isGreedy, "", "?") & ")(" & endRowPattern & "\s?)"
    Dim matches(): matches = GetRegExpSubMatches(template, regExpPattern) ' @dependency: 98.bas
    If GetArrLength(matches) > 0 Then ' @dependency: 2.bas
        GetBlockContents = matches
    Else
        GetBlockContents = Array()
    End If
End Function

Function GetStubsArr(ByVal template$) As Variant()
    ' Returns array of stubs from the template
    Dim stubsArr(): stubsArr = GetRegExpMatches(template, STUB_DELIMITER & "\??" & "[а-яА-Яa-zA-Z0-9_]*" & STUB_DELIMITER) ' @dependency: 60.bas
    stubsArr = GetUniqueArr(stubsArr) ' @dependency: 10.bas
    GetStubsArr = stubsArr
End Function

Function CleanResult(ByVal template$) As String
    ' Remove technical (IF, LOOP) and empty lines
    
    Dim key$, blockStartRow$, blockEndRow$
    Dim arr(), i&
    
    key = "[а-яА-Яa-zA-Z0-9_]+"
    
    blockStartRow = GetIfStartRow(key)
    blockEndRow = GetIfEndRow(key)
    arr = CombineArrays(arr, GetRegExpMatches(template, blockStartRow)) ' @dependency: 93.bas @dependency: 60.bas
    arr = CombineArrays(arr, GetRegExpMatches(template, blockEndRow)) ' @dependency: 93.bas @dependency: 60.bas
    
    blockStartRow = GetLoopStartRow(key)
    blockEndRow = GetLoopEndRow(key)
    arr = CombineArrays(arr, GetRegExpMatches(template, blockStartRow)) ' @dependency: 93.bas @dependency: 60.bas
    arr = CombineArrays(arr, GetRegExpMatches(template, blockEndRow)) ' @dependency: 93.bas @dependency: 60.bas
    
    If (Not arr) <> -1 Then
        For i = LBound(arr) To UBound(arr)
            template = Replace(template, arr(i), "")
        Next i
    End If
    
    template = Application.WorksheetFunction.Clean(template)
    
    CleanResult = template

End Function

Function InsertData2XmlTemplate(ByVal template$, ByRef stubsArr(), ByRef dataMap As Scripting.Dictionary) As String
    ' Fills in the template with data
    
    Dim stubIndex&
    For stubIndex = LBound(stubsArr) To UBound(stubsArr)
        
        Dim subStubsArr(), matches(), subDataMap As Scripting.Dictionary
        Dim blockStartRow$, blockEndRow$, blockContent$, rowToInsert$, rowsToInsert$, templateRow$
        Dim isCondition As Boolean, isLoop As Boolean, isValidData As Boolean, isOptional As Boolean
        
        Dim datakey$: datakey = stubsArr(stubIndex)
        datakey = RemoveStubDelimiters(datakey)
        
        isCondition = CBool(Left(datakey, 3) = "has")
        isLoop = CBool(Left(datakey, 4) = "each")
        
        Select Case (True)
            Case isCondition
                blockStartRow = GetIfStartRow(datakey)
                blockEndRow = GetIfEndRow(datakey)
            Case isLoop
                blockStartRow = GetLoopStartRow(datakey)
                blockEndRow = GetLoopEndRow(datakey)
        End Select
        
        blockContent = ""
        
        If isCondition Or isLoop Then
            matches = GetBlockContents(template, blockStartRow, blockEndRow)
            If GetArrLength(matches) > 0 Then ' @dependency: 2.bas
                blockContent = matches(0)(1)
            End If
        End If
        
        If dataMap.Exists(datakey) Then
        
            Select Case (True)
            
                Case isCondition
                    
                    isValidData = CBool(typeName(dataMap(datakey)) = "Dictionary" And dataMap(datakey).Count > 0)
                    
                    If isValidData Then
                    
                        Set subDataMap = dataMap(datakey)
                                        
                        subStubsArr = GetSubStubsArr(blockContent, subDataMap)
                        rowToInsert = InsertData2XmlTemplate(blockContent, subStubsArr, subDataMap)
                        Call InsertNewRow(template, rowToInsert, blockStartRow, blockContent)
                    
                    Else
                    
                        template = RemoveEmptyBlock(template, blockContent)
                        
                    End If
                    
                Case isLoop
                        
                    isValidData = CBool(IsArray(dataMap(datakey)) And UBound(dataMap(datakey)) > -1)
                
                    If isValidData Then
                    
                        Dim arrElem() As Scripting.Dictionary: arrElem = dataMap(datakey)
                               
                        Dim arrIndex&
                        For arrIndex = UBound(arrElem) To 0 Step -1
                            
                            Set subDataMap = arrElem(arrIndex)
                            
                            If Not subDataMap Is Nothing Then
                            
                                subStubsArr = GetSubStubsArr(blockContent, subDataMap)
                                rowToInsert = InsertData2XmlTemplate(blockContent, subStubsArr, subDataMap)
                                rowsToInsert = rowsToInsert & vbLf & rowToInsert
                                
                            End If
                            
                        Next arrIndex
                        
                        Call InsertNewRow(template, rowsToInsert, blockStartRow, blockContent)
                        rowsToInsert = ""
                        
                    Else
                    
                        template = RemoveEmptyBlock(template, blockContent)
                        
                    End If
                    
                Case Else
                
                    isOptional = IIf(Left(datakey, 1) = "?", True, False)
        
                    Dim val$: val = dataMap(datakey)
                    Dim key$: key = STUB_DELIMITER & datakey & STUB_DELIMITER
                    
                    If isOptional And val = "" Then
                        Dim arr() As String: arr = Split(key, "_")
                        Dim tag$: tag = arr(0)
                        Dim atr$: atr = arr(1)
                        tag = Right(tag, Len(tag) - 1)
                        atr = Left(atr, Len(atr) - 1)
                        key = " " & atr & "=""" & key & """"
                    End If
                    
                    template = Replace(template, key, val, , 1)
            
            End Select
            
        Else
            
            template = RemoveEmptyBlock(template, blockContent)

        End If
            
    Next stubIndex
        
    InsertData2XmlTemplate = template
    
End Function

Private Function GetSubStubsArr(ByVal template$, ByRef dataMap As Scripting.Dictionary) As Variant()
    ' Returns array of stubs from the template
    Dim stubsArr(), datakey
    For Each datakey In dataMap.Keys()
        Dim stub$: stub = GetFirstRegExpMatch(template, STUB_DELIMITER & "\??" & datakey & STUB_DELIMITER) ' @dependency: 62.bas
        Call AddToArr(stubsArr, stub) ' @dependency: 1.bas
    Next datakey
    GetSubStubsArr = stubsArr
End Function

Private Sub InsertNewRow(ByRef template$, ByRef rowToInsert$, ByVal afterRow$, ByVal blockContent$)
    ' Insert new row after ifStartRow
    rowToInsert = IIf(Right(rowToInsert, 1) = vbLf, Left(rowToInsert, Len(rowToInsert) - 1), rowToInsert)
    rowToInsert = IIf(Left(rowToInsert, 1) = vbLf, Right(rowToInsert, Len(rowToInsert) - 1), rowToInsert)
    Dim afterRowPos&: afterRowPos = InStr(1, template, afterRow) + Len(afterRow)
    Dim upperPart$: upperPart = Left(template, afterRowPos - 1)
    Dim lowerPart$: lowerPart = Mid(template, afterRowPos)
    lowerPart = Replace(lowerPart, blockContent, "", , 1)
    template = upperPart & vbLf & rowToInsert & lowerPart
End Sub

Private Function RemoveEmptyBlock(ByVal template$, ByVal block$) As String
    
    ' Delete template block if no data provided
    If block <> "" Then
        template = Replace(template, block, "", , 1)
    End If
    
    RemoveEmptyBlock = template

End Function