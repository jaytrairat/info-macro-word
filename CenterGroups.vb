Sub CenterGroups()
    Dim shp As shape
    Dim pageWidth As Double
    Dim pageMargin As Double
    
    pageWidth = ActiveDocument.PageSetup.pageWidth
    pageMarginLeft = ActiveDocument.PageSetup.LeftMargin
    pageMarginRight = ActiveDocument.PageSetup.RightMargin
    pageMargin = (pageWidth - pageMarginLeft - pageMarginRight) / 2
    
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoGroup Then
            shp.Left = pageMargin - (shp.Width / 2)
        End If
    Next shp
End Sub
