Sub HideAllTextBoxBorder()
    Dim shp As Shape
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoTextBox Then
            shp.Line.Visible = msoFalse
        End If
    Next shp