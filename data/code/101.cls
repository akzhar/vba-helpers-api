VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Enum EConfigColumns
    KeyColumn = 1
    ValueColumn
End Enum

Private Const ConfigTable As String = "ConfigTable"

Private Type TConfig
    Table  As ListObject
    Keys   As Range
    Values As Range
End Type

Private This As TConfig

Private Sub InitThis()
    Set This.Table = Me.Table
    Set This.Keys = Me.Keys
    Set This.Values = Me.Values
End Sub

Public Property Get Table() As ListObject
    Set Table = Me.ListObjects(ConfigTable)
End Property

Public Property Get Keys() As Range
    Set Keys = Me.Table.ListColumns(KeyColumn).DataBodyRange
End Property

Public Property Get Values() As Range
    Set Values = Me.Table.ListColumns(ValueColumn).DataBodyRange
End Property

Public Property Get Key(ByVal ThisKey As Variant) As Variant
    Call InitThis
    If This.Keys Is Nothing Then Key = "There is no data in " & ConfigTable: Exit Property

    Dim i As Long
    For i = 1 To This.Keys.Rows.Count
        If CStr(ThisKey) = CStr(This.Keys(i).Value) Then Key = This.Values(i).Value: Exit Property
    Next i
End Property

Public Property Let Key(ByVal ThisKey As Variant, ByVal RHS As Variant)
    Call InitThis
    If This.Keys Is Nothing Then This.Table.ListRows.Add

    Dim i As Long
    Do Until ThisKey = This.Keys(i).Value
         i = i + 1
         If i > This.Keys.Rows.Count Then This.Table.ListRows.Add: Exit Do
    Loop

    This.Keys(i).Value = ThisKey
    This.Values(i).Value = RHS
    Call DeleteEmptyConfigRows
End Property

Private Sub DeleteEmptyConfigRows()
    Call InitThis

    Dim i As Long
    For i = This.Keys.Count To 1 Step -1
         If (IsEmpty(This.Keys(i).Value) And IsEmpty(This.Values(i).Value)) _
         Or (This.Keys(i).Value = vbNullString And This.Values(i).Value = vbNullString) Then This.Table.ListRows(i).Delete
    Next i
End Sub

