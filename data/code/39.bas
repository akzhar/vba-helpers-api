Attribute VB_Name = "Helper39"
Option Explicit

Function GetSelectedRadioBtn(ByVal frameName$) As MSforms.OptionButton
    ' ф-ция находит и возвращает выбранную radio button по имени ее родителя frameName
    Dim ctrl As Control: Set ctrl = Nothing
    Dim opt As MSforms.OptionButton
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.Parent.Name = frameName Then
                Set opt = ctrl
                If opt.Value Then
                    Set GetSelectedRadioBtn = opt
                    Exit For
                End If
            End If
        End If
    Next ctrl
End Function