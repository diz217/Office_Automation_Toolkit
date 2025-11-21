Option Explicit

Public Sub LinkI()
    Dim sel As Selection
    Dim tr As TextRange
    Dim shp As Shape
    Dim idx As Long
    Dim sld As Slide
    Set sel = ActiveWindow.Selection
    Select Case sel.Type
        Case ppSelectionText
            Set tr = sel.TextRange
            Set shp = sel.ShapeRange(1)
        Case ppSelectionShapes
            Set shp = sel.ShapeRange(1)
            If shp.HasTextFrame Then
                If shp.TextFrame2.HasText Then
                    Set tr = shp.TextFrame.TextRange
                End If
            End If
 
        Case Else
            MsgBox "Select something", vbExclamation
            Exit Sub
    End Select
    If tr Is Nothing And shp Is Nothing Then
        MsgBox "None detected", vbExclamation
        Exit Sub
    End If

    idx = Val(InputBox("Link to Page Number?", "LinkI", CStr(ActiveWindow.View.Slide.SlideIndex + 2)))
    If idx < 1 Or idx > ActivePresentation.Slides.count Then
        MsgBox "Wrong number", vbExclamation
        Exit Sub
    End If
    Set sld = ActivePresentation.Slides(idx)
    If Not tr Is Nothing Then
        With tr.ActionSettings(ppMouseClick)
            .Action = ppActionHyperlink
            .Hyperlink.Address = ""
            .Hyperlink.SubAddress = sld.SlideID & "," & sld.SlideIndex & ","
        End With
    ElseIf Not shp Is Nothing Then
        With shp.ActionSettings(ppMouseClick)
            .Action = ppActionHyperlink
            .Hyperlink.Address = ""
            .Hyperlink.SubAddress = sld.SlideID & "," & sld.SlideIndex & ","
        End With
    End If
End Sub
 
Public Sub LinkII()
    Dim sel As Selection
    Dim shp As Shape
    Dim tr As TextRange
    Dim sld As Slide
    Dim idx As Long
    Dim shp2 As Shape
    Dim tr2 As TextRange
    Dim sldSrc As Slide
    Dim idxSrc As Long
    Dim newtext As String
    Set sel = ActiveWindow.Selection
    Select Case sel.Type
        Case ppSelectionText
            Set shp = sel.TextRange.Parent.Parent
        Case ppSelectionShapes
            Set shp = sel.ShapeRange(1)
        Case Else
            MsgBox "Select a textbox, text,or shape", vbExclamation
            Exit Sub
    End Select
    If Not shp.TextFrame.HasText Then
        MsgBox "No texts here", vbExclamation
        Exit Sub
    End If
    Set sldSrc = shp.Parent
    idxSrc = sldSrc.SlideIndex
    Set tr = shp.TextFrame.TextRange
    idx = Val(InputBox("Link to Page Number?", "LinkII", CStr(idxSrc + 2)))
    If idx < 1 Or idx > ActivePresentation.Slides.count Then
        MsgBox "Wrong number", vbExclamation
        Exit Sub
    End If
    Set sld = ActivePresentation.Slides(idx)
    shp.Copy
    Set shp2 = sld.Shapes.Paste(1)
    shp2.Left = shp.Left
    shp2.Top = shp.Top
    Set tr2 = shp2.TextFrame.TextRange
    newtext = InputBox("New name?", "LinkII-Name", tr.Text)
    tr2.Text = newtext
    With tr.ActionSettings(ppMouseClick)
        .Action = ppActionHyperlink
        .Hyperlink.Address = ""
        .Hyperlink.SubAddress = sld.SlideID & "," & sld.SlideIndex & ","
    End With
    With tr2.ActionSettings(ppMouseClick)
        .Action = ppActionHyperlink
        .Hyperlink.Address = ""
        .Hyperlink.SubAddress = sldSrc.SlideID & "," & sldSrc.SlideIndex & ","
    End With
End Sub
