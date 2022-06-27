Attribute VB_Name = "Helper79"
Option Explicit

Function PicturesExport(ByRef ws As Worksheet, ByVal pathToSave$)
    ' ф-ция экспортирует все изображения с листа в указанную папку

    Dim shp As Shape
    Dim counter&: counter =0

    Call TurnUpdatesOn(False) ' @(id 51)
    
    With ws
        For Each shp In .Shapes
            If shp.Type = msoPicture Then
                Charts.Add
                ActiveChart.Location xlLocationAsObject, .Name
                ActiveChart.ChartArea.Border.LineStyle = 0
                ActiveChart.ChartArea.Width = shp.Width
                ActiveChart.ChartArea.Height = shp.Height
                oShp.Copy
                ActiveChart.ChartArea.Select
                ActiveChart.Paste
                counter = counter + 1
                .ChartObjects(1).Chart.Export _
                  FileName:=pathToSave & Application.PathSeparator & counter & ".jpg", _
                  FilterName:="jpg"
                .ChartObjects(1).Delete
            End If
        Next shp
    End With
    
    Call TurnUpdatesOn(True) ' @(id 51)

    MsgBox "Сохранено " & counter & " изображений", vbInformation

End Function