Public Sub PasteObject()
    Dim sel As Selection
    Dim shp As Shape
    Dim sld As Slide
    Dim bleft As Single, btop As Single, bwidth As Single, bheight As Single
    Dim newshp As Shape
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        If sel.Type <> ppSelectionText Then
            MsgBox "Not a shape", vbExclamation
            Exit Sub
        End If
    End If
    Set shp = sel.ShapeRange(1)
    bleft = shp.Left
    btop = shp.Top
    bwidth = shp.Width
    bheight = shp.Height
    startIndex = ActiveWindow.View.Slide.SlideIndex
    shp.Copy
    For Each sld In ActivePresentation.Slides
        If sld.SlideIndex > startIndex Then
            sld.Shapes.PasteSpecial
            Set newshp = sld.Shapes(sld.Shapes.count)
            With newshp
                .Left = bleft
                .Top = btop
                .Width = bwidth
                .Height = bheight
            End With
        End If
    Next sld
End Sub
 
Public Sub PasteObjectII()
    Dim sel As Selection
    Dim shp As Shape
    Dim sld As Slide
    Dim bleft As Single, btop As Single, bwidth As Single, bheight As Single
    Dim newshp As Shape
    Dim step As Long
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        If sel.Type <> ppSelectionText Then
            MsgBox "Not a shape", vbExclamation
            Exit Sub
        End If
    End If
    Set shp = sel.ShapeRange(1)
    bleft = shp.Left
    btop = shp.Top
    bwidth = shp.Width
    bheight = shp.Height
    startIndex = ActiveWindow.View.Slide.SlideIndex
    step = Val(InputBox("Interval:", , "2"))
    If step < 1 Then step = 2
    shp.Copy
    For Each sld In ActivePresentation.Slides
        If sld.SlideIndex > startIndex Then
            If (sld.SlideIndex - startIndex) Mod step = 0 Then
                sld.Shapes.PasteSpecial
                Set newshp = sld.Shapes(sld.Shapes.count)
                With newshp
                    .Left = bleft
                    .Top = btop
                    .Width = bwidth
                    .Height = bheight
                End With
            End If
        End If
    Next sld
End Sub
