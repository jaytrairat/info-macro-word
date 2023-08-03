Sub CenterGroupsAndImages()
    Dim shp As shape
    Dim pageWidth As Double
    Dim pageMarginLeft As Double
    Dim pageMarginRight As Double
    Dim pageMargin As Double

    pageWidth = ActiveDocument.PageSetup.pageWidth
    pageHeight = ActiveDocument.PageSetup.pageHeight
    pageMarginTop = ActiveDocument.PageSetup.TopMargin
    pageMarginBottom = ActiveDocument.PageSetup.BottomMargin
    pageMarginLeft = ActiveDocument.PageSetup.LeftMargin
    pageMarginRight = ActiveDocument.PageSetup.RightMargin

    contentHeight = pageHeight - pageMarginTop - pageMarginBottom
    contentWidth = pageWidth - pageMarginLeft - pageMarginRight
    centerOfPage = contentWidth / 2

    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoGroup Or shp.Type = msoPicture Then
            shp.Left = centerOfPage - (shp.Width / 2)
        ElseIf shp.Type = msoTextBox Then
            If InStr(1, shp.TextFrame.TextRange.Text, "...") > 0 Then
                shp.Left = contentWidth - (shp.Width)
                shp.Top = contentHeight
            Else
                shp.Left = centerOfPage - (shp.Width / 2)
            End If
        End If
    Next shp

    MsgBox "Centered all shapes"
End Sub
