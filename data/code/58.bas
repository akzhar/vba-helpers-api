Attribute VB_Name = "VbaHelper_RefreshPQ"
Option Explicit

Function RefreshPQ(ByVal queryName$)
    ' Refreshes PowerQuery by name of the query

    Dim con As WorkbookConnection

    For Each con In ThisWorkbook.Connections
        If (con.Name = "Query - " & queryName) Then
            With ThisWorkbook.Connections(con.Name).OLEDBConnection
                .BackgroundQuery = True
                .Refresh
                ' Waiting when Power Query refresh is completed
                Call WaitTillRefreshComplete
            End With
        End If
    Next con
End Function

Private Sub WaitTillRefreshComplete()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)

    Dim b1 As Boolean: b1 = ws.Range("Query_1_Range_Name").ListObject.QueryTable.Refreshing
    Dim b2 As Boolean: b2 = ws.Range("Query_2_Range_Name").ListObject.QueryTable.Refreshing

    If b1 Or b2 Then
        Call Application.OnTime(Now + TimeValue("00:00:01"), "WaitTillRefreshComplete")
    Else
        Call Application.Run("DoAfterRefreshComplete")
    End If
    
End Sub

Private Sub DoAfterRefreshComplete()
    
    MsgBox "Query has been refreshed", vbInformation

End Sub