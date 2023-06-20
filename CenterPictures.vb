Sub CenterPictures()
   Dim shp As shape
    pageWidth = ActiveDocument.PageSetup.pageWidth
    pageMarginLeft = ActiveDocument.PageSetup.LeftMargin
    pageMarginRight = ActiveDocument.PageSetup.RightMargin
    pageMargin = (pageWidth - pageMarginLeft - pageMarginRight) / 2
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoPicture Then
            shp.Select
            shp.Left = pageMargin - (shp.Width / 2)
        End If
    Next shp
End Sub
