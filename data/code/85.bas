Attribute VB_Name = "Helper85"
Option Explicit

Function SetComment(ByRef rng As Range, ByVal flag As Boolean, Optional ByVal comment$ = "", Optional isVisible As Boolean = False)
    
    If flag Then
    
        With rng
            If CBool(.comment Is Nothing) Then
                .AddComment
            End If
            .comment.Visible = isVisible
            .comment.Text Text:=comment
            .comment.Shape.TextFrame.Characters.Font.Size = 12
        End With
        
        Call FitComments ' @(id 52)
    
    Else
    
        rng.ClearComments
    
    End If
    
End Function