Attribute VB_Name = "Helper58"
Option Explicit

Function RefreshPQ(ByVal queryName$)
    ' ф-ция обновляет Power Query запрос с именем queryName
    Dim con As WorkbookConnection
    For Each con In ActiveWorkbook.Connections
        If (con.Name = "Query - " & queryName) Then
            With ActiveWorkbook.Connections(con.Name).OLEDBConnection
		            ' ожидание обновления Power Query
                .BackgroundQuery = False
                .Refresh
            End With
        End If
    Next
End Function