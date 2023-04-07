Attribute VB_Name = "VbaHelper_SetNote"
Option Explicit

Function SetNote(ByRef rng As Range, ByVal flag As Boolean, Optional ByVal comment$ = "", Optional isVisible As Boolean = False)
    ' Sets note in the specified cell
    If flag Then
        With rng
            If CBool(.comment Is Nothing) Then
                .AddComment
            End If
            .comment.Visible = isVisible
            .comment.Text Text:=comment
            .comment.Shape.TextFrame.Characters.Font.Size = 12
            .comment.Shape.TextFrame.AutoSize = True
        End With
    Else
        rng.ClearComments
    End If
End Function