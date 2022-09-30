Attribute VB_Name = "Helper58"
Option Explicit

Function RefreshPQ(ByVal queryName$)
    ' Refreshes PowerQuery by name of the query

    Dim con As WorkbookConnection

    For Each con In ActiveWorkbook.Connections
        If (con.Name = "Query - " & queryName) Then
            With ActiveWorkbook.Connections(con.Name).OLEDBConnection
                .BackgroundQuery = True
                .Refresh
                ' waiting when Power Query refresh is complete
                Call WaitRefreshComplete
            End With
        End If
    Next
End Function

Private Sub WaitRefreshComplete()

    Dim t: t = TimeValue("00:00:01")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)

    Dim b1 As Boolean: b1 = ws.Range("Query1 Name").ListObject.QueryTable.Refreshing
    Dim b2 As Boolean: b2 = ws.Range("Query2 Name").ListObject.QueryTable.Refreshing

    If b1 Or b2 Then
        Call Application.OnTime(Now + t, "WaitRefreshComplete")
    Else
        Call Application.Run("DoAfterRefreshComplete")
    End If
    
End Sub

Private Sub DoAfterRefreshComplete()
    
    MsgBox "Query has been refreshed"

End Sub