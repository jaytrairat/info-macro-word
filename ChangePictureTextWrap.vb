Sub ChangePictureTextWrap()
    
    Dim shp As Shape
    
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoPicture Then
            shp.Select
            Selection.ShapeRange.WrapFormat.Type = wdWrapTopBottom
        End If
    Next shp
    

End Sub
