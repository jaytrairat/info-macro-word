Sub CenterGroupsAndImages()
    Dim shp As shape
    Dim contentWidth As Double
    Dim centerOfPage As Double
    
    contentWidth = ActiveDocument.PageSetup.pageWidth - ActiveDocument.PageSetup.LeftMargin - ActiveDocument.PageSetup.RightMargin
    centerOfPage = contentWidth / 2

    For Each shp In ActiveDocument.Shapes
        Select Case shp.Type
            Case msoGroup, msoPicture
                shp.Left = centerOfPage - (shp.Width / 2)
            Case msoTextBox
                If shp.TextFrame.TextRange.Paragraphs.Count > 1 Then
                    shp.Left = contentWidth - shp.Width
                ElseIf InStr(shp.TextFrame.TextRange.Text, "/") <> 0 Then
                    shp.Left = contentWidth - shp.Width
                    shp.Top = ActiveDocument.PageSetup.pageHeight - ActiveDocument.PageSetup.BottomMargin
                Else
                    shp.Left = centerOfPage - (shp.Width / 2)
                End If
        End Select
    Next shp

    MsgBox "Centered all shapes"
End Sub
