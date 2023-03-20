Sub CenterPictures()
    Dim shp As Shape
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoPicture Then
            shp.Select
            Set shpRange = Selection.ShapeRange
            shpRange.Align msoAlignCenters, True

        End If
    Next shp
End Sub