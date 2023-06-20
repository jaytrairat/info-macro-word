Sub CenterGroupsAndImages()
    Dim shp As shape
    Dim pageWidth As Double
    Dim pageMarginLeft As Double
    Dim pageMarginRight As Double
    Dim pageMargin As Double
    
    pageWidth = ActiveDocument.PageSetup.pageWidth
    pageMarginLeft = ActiveDocument.PageSetup.LeftMargin
    pageMarginRight = ActiveDocument.PageSetup.RightMargin
    pageMargin = (pageWidth - pageMarginLeft - pageMarginRight) / 2
    
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoGroup Or shp.Type = msoPicture Or shp.Type = msoTextBox Then
            shp.Left = pageMargin - (shp.Width / 2)
        End If
    Next shp
    
    MsgBox "Centered all shapes"
End Sub
