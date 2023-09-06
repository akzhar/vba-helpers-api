VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class to manipulate tables

Private tblHeadersDict As Dictionary
Private tblObj As ListObject

Private Sub Class_Initialize()
    Set tblHeadersDict = New Dictionary
End Sub
    
Private Sub Class_Terminate()
    Call Reset
End Sub

' Public Properties

Public Property Get TableName() As String

    TableName = tblObj.Name
    
End Property

Public Property Get Range() As Range

    Set Range = tblObj.Range
    
End Property

Public Property Get Sheet() As Worksheet

    Set Sheet = tblObj.Parent
    
End Property

Public Property Get LastColumnNum() As Long

    LastColumnNum = Me.HeadersRow(Me.HeadersRow.Columns.Count).Column
    
End Property

Public Property Get HeadersRow() As Range

    Set HeadersRow = tblObj.HeaderRowRange
    
End Property

Public Property Get Rows() As Range

    Set Rows = tblObj.DataBodyRange.Rows
    
End Property

Public Property Get RowsVisible() As Range

    Dim filteredRows As Range
    Dim rowRng As Range

    For Each rowRng In Me.Rows()
        
        If rowRng.EntireRow.Hidden = False Then

            If filteredRows Is Nothing Then
                Set filteredRows = rowRng
            Else
                Set filteredRows = Union(filteredRows, rowRng)
            End If

        End If
        
    Next rowRng

    Set RowsVisible = filteredRows.Rows

End Property

' Private Properties

Private Property Get ColumnIndexByHeader() As Dictionary

    Set ColumnIndexByHeader = tblHeadersDict

End Property

' Private methods

Private Function IsHeaderExists(ByVal headerValue$) As Boolean
    
    IsHeaderExists = ColumnIndexByHeader.Exists(headerValue)
    
End Function

Private Function GetRowIndex(ByVal headerValue$, ByVal lookupValue$) As Long

    GetRowIndex = -1

    If Not IsHeaderExists(headerValue) Then Exit Function
    
    Dim foundCell As Range
    Dim rowIndex&, colIndex&
    
    colIndex = ColumnIndexByHeader(headerValue)
    
    ' Attempt to find value in table
    
    On Error Resume Next
    Set foundCell = tblObj.DataBodyRange.Columns(colIndex).Find(lookupValue, LookAt:=xlWhole)
    On Error GoTo 0

    If Not foundCell Is Nothing Then
        rowIndex = tblObj.ListRows(foundCell.Row - Me.HeadersRow.Row).Index
        GetRowIndex = rowIndex
    End If
    
End Function

' Public methods

Public Sub Reset()

    Set tblHeadersDict = Nothing
    tblObj.Unlist
    Set tblObj = Nothing
    
End Sub

Public Sub Create(ByRef sourceRng As Range)
    
    Dim tblWs As Worksheet: Set tblWs = sourceRng.Parent

    ' Get / create table object
    
    Dim tblName$
    
    tblName = Replace(Replace(tblWs.Name & " " & "ListObj", " ", "_"), "-", "_")
    tblName = IIf(IsNumeric(Left(tblName, 1)), "_" & tblName, tblName)
    
    If Not Utils.IsNamedRangeExists(tblWs, tblName) Then ' @dependency: 118.bas
        tblWs.AutoFilterMode = False
        tblWs.ListObjects.Add(SourceType:=xlSrcRange, source:=sourceRng, XlListObjectHasHeaders:=xlYes).Name = tblName
        tblWs.ListObjects(tblName).TableStyle = ""
        tblWs.ListObjects(tblName).ShowAutoFilterDropDown = False
    End If
    
    Set tblObj = tblWs.ListObjects(tblName)

    ' Save headers values and column numbers in Dictionary
    
    Dim cell As Range
    Dim colsCounter&: colsCounter = 1
    
    For Each cell In Me.HeadersRow
        tblHeadersDict.Add Trim(CStr(cell.value)), colsCounter
        colsCounter = colsCounter + 1
    Next cell
        
End Sub

Public Function GetCell(ByVal colName$, Optional ByVal rowFilter$, Optional rowIndex&) As Range

    Set GetCell = Nothing
    
    If rowFilter = "" And rowIndex = 0 Then Exit Function
    
    If Not IsHeaderExists(colName) Then Exit Function
    
    Dim colIndex&: colIndex = ColumnIndexByHeader(colName)
        
    If rowIndex = 0 Then
        Dim arr() As String: arr = Split(rowFilter, "::")
        rowIndex = GetRowIndex(headerValue:=arr(0), lookupValue:=arr(1))
    End If
    
    If rowIndex = -1 Then Exit Function
    
    Set GetCell = tblObj.DataBodyRange(rowIndex, colIndex)
    
End Function

Public Sub SetValue(ByVal value$, ByVal setCol$, Optional ByVal rowFilter$, Optional rowIndex&)
    
    If rowFilter = "" And rowIndex = 0 Then
        Exit Sub
    End If
    
    If Not IsHeaderExists(setCol) Then Exit Sub
    
    Dim colIndex&: colIndex = ColumnIndexByHeader(setCol)
        
    If rowIndex = 0 Then
        Dim arr() As String: arr = Split(rowFilter, "::")
        rowIndex = GetRowIndex(headerValue:=arr(0), lookupValue:=arr(1))
    End If
    
    If rowIndex = -1 Then Exit Sub
    
    tblObj.DataBodyRange(rowIndex, colIndex).value = value
    
End Sub

Public Sub ClearFilters()

    ' Clear any existing filter
    If tblObj.ShowAutoFilter Then
        tblObj.ShowAutoFilter = False
    End If
    
End Sub

Public Sub SetFilters(ByVal filters$)
    
    Call ClearFilters
    
    Dim filtersArr() As String: filtersArr = Split(filters, ";")
    
    Dim i&
    For i = LBound(filtersArr) To UBound(filtersArr)
        
        Dim arr() As String: arr = Split(filtersArr(i), "::")
        Dim colIndex&: colIndex = ColumnIndexByHeader(arr(0))
        ' Set filter
        tblObj.Range.AutoFilter Field:=colIndex, Criteria1:=arr(1)
    
    Next i
    
End Sub

Public Function AddRow() As Range

    Dim newRow As ListRow: Set newRow = tblObj.ListRows.Add(AlwaysInsert:=True)
    Set AddRow = newRow.Range
    
End Function

Public Function AddColumn(ByVal colName$) As Range

    Dim newCol As ListColumn: Set newCol = tblObj.ListColumns.Add()
    newCol.Name = Trim(colName)
    
    ' Add new header value and column number in Dictionary
    
    tblHeadersDict.Add Trim(colName), Me.LastColumnNum
    
    Set AddColumn = newCol.Range
    
End Function

Public Function IsColumnExists(ByVal colName$) As Boolean

    IsColumnExists = IsHeaderExists(colName)
    
End Function

Public Function DeleteRows(ByVal colName$, ByVal criteria$) As Long
    ' Delete rows matches specified criteria (value or color)
    ' Returns how many rows were deleted

    DeleteRows = 0

    Dim criteriaSign$, criteriaValue$
    
    Select Case True
        Case criteria Like "=*":
            criteriaSign = "="
        Case criteria Like "<>*":
            criteriaSign = "<>"
        Case criteria Like ">*":
            criteriaSign = ">"
        Case criteria Like "<*":
            criteriaSign = "<"
        Case criteria Like ">=*":
            criteriaSign = ">="
        Case criteria Like "<=*":
            criteriaSign = "<="
        Case criteria Like "color:*":
            criteriaSign = "color:"
        Case Else:
            criteriaSign = "?"
    End Select
    
    If criteriaSign = "?" Then Exit Function
    
    criteriaValue = Replace(criteria, criteriaSign, "")

    Dim i&, deletedRows&, tblRow As ListRow, colCell As Range
    
    For i = tblObj.ListRows.Count To 1 Step -1
    
        Set tblRow = tblObj.ListRows(i)
        Set colCell = Intersect(tblRow.Range, tblObj.ListColumns(colName).Range)
        
        ' @dependency: 36.bas
        
        If _
            IIf( _
                criteriaSign = "color:", _
                Utils.IsColoredLike(colCell, criteriaValue), _
                Application.Evaluate(colCell.value & criteriaSign & criteriaValue) _
            ) _
        Then
            tblRow.Delete
            deletedRows = deletedRows + 1
        End If
        
    Next i
    
    DeleteRows = deletedRows
    
End Function

Public Function GetColumn(ByVal colName$, Optional ByVal includeHeader As Boolean = True) As Range
    
    Set GetColumn = Nothing
    
    Dim colIndex&: colIndex = ColumnIndexByHeader(colName)
    
    If colIndex = 0 Then Exit Function
    
    Dim colRng As Range: Set colRng = tblObj.ListColumns(colIndex).Range
    
    If Not includeHeader Then
        
        Set colRng = colRng.Offset(1, 0).Resize(colRng.Rows.Count - 1, 1)
        
    End If
    
    Set GetColumn = colRng
    
End Function

Public Function GetColumnValues(ByVal colName$, Optional ByVal onlyUnique As Boolean = False) As Variant()
    ' Returns array of values in specified column of table
    
    Dim values()
    
    Dim valuesRng As Range: Set valuesRng = Me.GetColumn(colName:=colName, includeHeader:=False)
    
    If Not valuesRng Is Nothing Then
            
        values = Utils.Rng2Array(rng:=valuesRng) ' @dependency: 3.bas
        
        If onlyUnique Then
        
            values = Utils.GetUniqueArr(arr:=values) ' @dependency: 10.bas
            
        End If
        
    End If
    
    GetColumnValues = values
    
End Function
