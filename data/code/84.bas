Attribute VB_Name = "Helper84"
Option Explicit

Public Const VAR_DELIMITER$ = "%"

Private Function GetRegExpPattern(ByVal startRow$, ByVal endRow$) As String
    ' Ф-ция возвращает паттерн для регулярки
    GetRegExpPattern = "(?:" & startRow & "$\s)([\S\s]+?)(?:" & endRow & "$\s?)"
End Function

Private Function GetPlugsArr(ByVal template$) As Variant()
    ' Ф-ция возвращает массив уникальных заглушек из шаблона
    Dim plugsArr(): plugsArr = GetRegExpMatches(template, VAR_DELIMITER & "\w*" & VAR_DELIMITER) ' @(id 60)
    plugsArr = GetUniqueArr(plugsArr) ' @(id 10)
    GetPlugsArr = plugsArr
End Function

Private Function GetBlockContent(ByVal template$, ByVal startRow$, ByVal endRow$) As String
    ' Ф-ция возвращает контент блока между строками startRow и endRow
    Dim matches(): matches = GetRegExpMatches(template, GetRegExpPattern(startRow, endRow)) ' @(id 60)
    GetBlockContent = ""
    'If Not matches Is Nothing Then
    If GetArrLength(matches) > 0 Then ' @(id 2)
        GetBlockContent = SliceString(template, InStr(1, template, startRow) + Len(startRow), InStr(1, template, endRow) - 1) ' @(id 72)
    End If
End Function

Private Sub InsertNewRow(ByRef template$, ByRef rowToInsert$, ByVal afterRow$)
    ' Вставка новой строки после строки ifStartRow
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

Private Function FillTemplateWithData(ByVal template$, ByRef plugsArr(), ByRef dataMap As Scripting.Dictionary) As String
    ' Ф-ция подставляет в template данные из map
    ' Значения из map вставляются на места обозначенные как %key% в шаблоне
    ' Если эл-т map является массивом, состоящим из map, то для каждого map из этого массива формируется строка по шаблону
    ' Шаблон строки должен быть расположен в исходном шаблоне между строками loopStartRow и loopEndRow
    
    Dim plugIndex&
    For plugIndex = LBound(plugsArr) To UBound(plugsArr)
        
        Dim subPlugsArr()
        Dim dataObj As Scripting.Dictionary
        Dim blockStartRow$, blockEndRow$, rowToInsert$, templateRow$, blockContent$
        Dim isCondition As Boolean, isLoop As Boolean, isValidData As Boolean
        
        Dim dataKey$: dataKey = plugsArr(plugIndex)
        dataKey = Replace(dataKey, VAR_DELIMITER, "")
        
        blockStartRow = GetIfStartRow(dataKey)
        blockEndRow = GetIfEndRow(dataKey)
        blockContent = GetBlockContent(template, blockStartRow, blockEndRow)
        
        isCondition = CBool(blockContent <> "" And InStr(1, blockContent, VAR_DELIMITER) > 0)
        isValidData = CBool(TypeName(dataMap(dataKey)) = "Dictionary")
            
        If isCondition Then
            
            If Not IsEmpty(dataMap(dataKey)) And isValidData Then
            
                Set dataObj = dataMap(dataKey)
                                
                subPlugsArr = GetPlugsArr(blockContent)
                rowToInsert = FillTemplateWithData(blockContent, subPlugsArr, dataObj)
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
                        subPlugsArr = GetPlugsArr(blockContent)
                        rowToInsert = FillTemplateWithData(blockContent, subPlugsArr, dataObj)
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
        
    Next plugIndex
        
    FillTemplateWithData = template
    
End Function