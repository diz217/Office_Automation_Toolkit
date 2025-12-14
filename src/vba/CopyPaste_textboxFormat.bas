Public Type TextBoxFormat
    FontName As String
    FontSize As Single
    FontBold As MsoTriState
    FontItalic As MsoTriState
    FontUnderline As MsoTriState
    FontColor As Long
    HighlightColor As Long
    FillVisible As MsoTriState
    FillForeColor As Long
    FillTransparency As Single
    LineVisible As MsoTriState
    LineForeColor As Long
    LineWeight As Single
    LineType As MsoLineDashStyle
    TextAlign As Long
    VerticalAlign As Long
    MarginLeft As Single
    MarginRight As Single
    MarginTop As Single
    MarginBottom As Single
End Type
 
Public StoredTB As TextBoxFormat
 
Public Sub CopyTBFormat()
    Dim shp As Shape
    If ActiveWindow.Selection.Type = ppSelectionText Or ActiveWindow.Selection.Type = ppSelectionShapes Then
        Set shp = ActiveWindow.Selection.ShapeRange(1)
    Else
        MsgBox "Please select a textbox", vbExclamation
        Exit Sub
    End If
    If shp.HasTextFrame = msoFalse Or shp.TextFrame2.HasText = msoFalse Then
        MsgBox "This shape is not textbox", vbExclamation
        Exit Sub
    End If
    Dim tr As TextRange2
    Set tr = shp.TextFrame2.TextRange
    'font
    StoredTB.FontName = tr.Font.Name
    StoredTB.FontSize = tr.Font.Size
    StoredTB.FontBold = tr.Font.Bold
    StoredTB.FontItalic = tr.Font.Italic
    StoredTB.FontUnderline = shp.TextFrame.TextRange.Font.Underline
    StoredTB.FontColor = tr.Font.Fill.ForeColor.RGB
    StoredTB.HighlightColor = tr.Font.Highlight.RGB
    'fill
    StoredTB.FillVisible = shp.Fill.Visible
    StoredTB.FillForeColor = shp.Fill.ForeColor.RGB
    StoredTB.FillTransparency = shp.Fill.Transparency
    'frame
    StoredTB.LineVisible = shp.Line.Visible
    If shp.Line.Visible = msoTrue Then
        StoredTB.LineForeColor = shp.Line.ForeColor.RGB
        StoredTB.LineWeight = shp.Line.Weight
        StoredTB.LineType = shp.Line.DashStyle
    Else
        StoredTB.LineWeight = 0
    End If
    'alignment
    StoredTB.TextAlign = shp.TextFrame2.TextRange.ParagraphFormat.Alignment
    StoredTB.VerticalAlign = shp.TextFrame2.VerticalAnchor
    StoredTB.MarginLeft = shp.TextFrame2.MarginLeft
    StoredTB.MarginRight = shp.TextFrame2.MarginRight
    StoredTB.MarginTop = shp.TextFrame2.MarginTop
    StoredTB.MarginBottom = shp.TextFrame2.MarginBottom
End Sub
 
Public Sub PasteTBFomrat()
    Dim shp As Shape
    Dim rng As ShapeRange
    If ActiveWindow.Selection.Type = ppSelectionText Or ActiveWindow.Selection.Type = ppSelectionShapes Then
        'Set shp = ActiveWindow.Selection.ShapeRange(1)
        Set rng = ActiveWindow.Selection.ShapeRange
    Else
        MsgBox "Please select one or more textboxs", vbExclamation
        Exit Sub
    End If
    For Each shp In rng
        If shp.HasTextFrame = msoFalse Or shp.TextFrame2.HasText = msoFalse Then
            MsgBox "This shape is not textbox", vbExclamation
            Exit Sub
        End If
        Dim tr As TextRange2
        Set tr = shp.TextFrame2.TextRange
        'font
        tr.Font.Name = StoredTB.FontName
        tr.Font.Size = StoredTB.FontSize
        tr.Font.Bold = StoredTB.FontBold
        tr.Font.Italic = StoredTB.FontItalic
        shp.TextFrame.TextRange.Font.Underline = StoredTB.FontUnderline
        tr.Font.Fill.ForeColor.RGB = StoredTB.FontColor
        If StoredTB.HighlightColor <> 0 Then
            tr.Font.Highlight.RGB = StoredTB.HighlightColor
        Else
            If StoredTB.FillVisible <> 0 Then
                tr.Font.Highlight.RGB = StoredTB.FillForeColor
            End If
        End If
        'fill
        shp.Fill.Visible = StoredTB.FillVisible
        shp.Fill.ForeColor.RGB = StoredTB.FillForeColor
        shp.Fill.Transparency = StoredTB.FillTransparency
        'frame
        shp.Line.Visible = StoredTB.LineVisible
        If StoredTB.LineVisible = msoTrue Then
            shp.Line.ForeColor.RGB = StoredTB.LineForeColor
            shp.Line.Weight = StoredTB.LineWeight
            shp.Line.DashStyle = StoredTB.LineType
        End If
        'alignment
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = StoredTB.TextAlign
        shp.TextFrame2.VerticalAnchor = StoredTB.VerticalAlign
        shp.TextFrame2.MarginLeft = StoredTB.MarginLeft
        shp.TextFrame2.MarginRight = StoredTB.MarginRight
        shp.TextFrame2.MarginTop = StoredTB.MarginTop
        shp.TextFrame2.MarginBottom = StoredTB.MarginBottom
    Next shp
End Sub
