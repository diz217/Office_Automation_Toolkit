Option Explicit
 
Public Type CopyShapeInfo
 
   HasData As Boolean
  
   Leftpos As Single
   TopPos As Single
   WidthVal As Single
   HeightVal As Single
  
   Rotation As Single
   HFlip As MsoTriState
   VFlip As MsoTriState
  
   HasFill As Boolean
   FillVisible As MsoTriState
   FillRGB As Long
  
   HasLine As Boolean
   LineVisible As MsoTriState
   LineRGB As Long
   LineWeight As Single
   LineType As msoLineDashStyle
  
   IsPicture As Boolean
   CropLeft As Single
   CropTop As Single
   CropRight As Single
   CropBottom As Single
End Type
 
Public gInfoA As CopyShapeInfo
Public gInfoB As CopyShapeInfo
Public gInfoC As CopyShapeInfo


Public Sub CopyShapeFormatA()
    Dim sel As Selection
    Dim shp As Shape
    Dim tmp As Shape
    Dim sld As Slide
    Dim picCount As Long
    Set sel = ActiveWindow.Selection
    Set sld = ActiveWindow.View.Slide
    If sel.Type = ppSelectionShapes Then
        If sel.ShapeRange.count <> 1 Then
            MsgBox "which one", vbExclamation
            Exit Sub
        End If
        Set shp = sel.ShapeRange(1)
    Else
        picCount = 0
        For Each tmp In sld.Shapes
            If tmp.Type = msoPicture Or tmp.Type = msoLinkedPicture Then
                Set shp = tmp
                picCount = picCount + 1
            End If
        Next tmp
        If picCount <> 1 Then
            MsgBox "select a shape", vbExclamation
            Exit Sub
        End If
    End If
    ' geometry
    gInfoA.Leftpos = shp.Left
    gInfoA.TopPos = shp.Top
    gInfoA.WidthVal = shp.Width
    gInfoA.HeightVal = shp.Height
    gInfoA.Rotation = shp.Rotation
    gInfoA.HFlip = shp.HorizontalFlip
    gInfoA.VFlip = shp.VerticalFlip
    ' fill
    gInfoA.HasFill = True
    gInfoA.FillVisible = shp.Fill.Visible
    If shp.Fill.Visible = msoTrue Then
        On Error Resume Next
        gInfoA.FillRGB = shp.Fill.ForeColor.RGB
        On Error GoTo 0
    End If
    ' frames
    gInfoA.HasLine = True
    gInfoA.LineVisible = shp.Line.Visible
    If shp.Line.Visible = msoTrue Then
        On Error Resume Next
        gInfoA.LineRGB = shp.Line.ForeColor.RGB
        gInfoA.LineWeight = shp.Line.Weight
        gInfoA.LineType = shp.Line.DashStyle
        On Error GoTo 0
    End If
    ' crop
    gInfoA.IsPicture = (shp.Type = msoPicture Or shp.Type = msoLinkedPicture)
    If gInfoA.IsPicture Then
        With shp.PictureFormat
            gInfoA.CropLeft = .CropLeft
            gInfoA.CropTop = .CropTop
            gInfoA.CropRight = .CropRight
            gInfoA.CropBottom = .CropBottom
        End With
    End If
    gInfoA.HasData = True
End Sub

Public Sub CopyShapeFormatB()
    Dim sel As Selection
    Dim shp As Shape
    Dim tmp As Shape
    Dim sld As Slide
    Dim picCount As Long
    Set sel = ActiveWindow.Selection
    Set sld = ActiveWindow.View.Slide
    If sel.Type = ppSelectionShapes Then
        If sel.ShapeRange.count <> 1 Then
            MsgBox "which one", vbExclamation
            Exit Sub
        End If
        Set shp = sel.ShapeRange(1)
    Else
        picCount = 0
        For Each tmp In sld.Shapes
            If tmp.Type = msoPicture Or tmp.Type = msoLinkedPicture Then
                Set shp = tmp
                picCount = picCount + 1
            End If
        Next tmp
        If picCount <> 1 Then
            MsgBox "select a shape", vbExclamation
            Exit Sub
        End If
    End If
    ' geometry
    gInfoB.Leftpos = shp.Left
    gInfoB.TopPos = shp.Top
    gInfoB.WidthVal = shp.Width
    gInfoB.HeightVal = shp.Height
    gInfoB.Rotation = shp.Rotation
    gInfoB.HFlip = shp.HorizontalFlip
    gInfoB.VFlip = shp.VerticalFlip
    ' fill
    gInfoB.HasFill = True
    gInfoB.FillVisible = shp.Fill.Visible
    If shp.Fill.Visible = msoTrue Then
        On Error Resume Next
        gInfoB.FillRGB = shp.Fill.ForeColor.RGB
        On Error GoTo 0
    End If
    ' frames
    gInfoB.HasLine = True
    gInfoB.LineVisible = shp.Line.Visible
    If shp.Line.Visible = msoTrue Then
        On Error Resume Next
        gInfoB.LineRGB = shp.Line.ForeColor.RGB
        gInfoB.LineWeight = shp.Line.Weight
        gInfoB.LineType = shp.Line.DashStyle
        On Error GoTo 0
    End If
    ' crop
    gInfoB.IsPicture = (shp.Type = msoPicture Or shp.Type = msoLinkedPicture)
    If gInfoB.IsPicture Then
        With shp.PictureFormat
            gInfoB.CropLeft = .CropLeft
            gInfoB.CropTop = .CropTop
            gInfoB.CropRight = .CropRight
            gInfoB.CropBottom = .CropBottom
        End With
    End If
    gInfoB.HasData = True
End Sub

Public Sub CopyShapeFormatC()
    Dim sel As Selection
    Dim shp As Shape
    Dim tmp As Shape
    Dim sld As Slide
    Dim picCount As Long
    Set sel = ActiveWindow.Selection
    Set sld = ActiveWindow.View.Slide
    If sel.Type = ppSelectionShapes Then
        If sel.ShapeRange.count <> 1 Then
            MsgBox "which one", vbExclamation
            Exit Sub
        End If
        Set shp = sel.ShapeRange(1)
    Else
        picCount = 0
        For Each tmp In sld.Shapes
            If tmp.Type = msoPicture Or tmp.Type = msoLinkedPicture Then
                Set shp = tmp
                picCount = picCount + 1
            End If
        Next tmp
        If picCount <> 1 Then
            MsgBox "select a shape", vbExclamation
            Exit Sub
        End If
    End If
    ' geometry
    gInfoC.Leftpos = shp.Left
    gInfoC.TopPos = shp.Top
    gInfoC.WidthVal = shp.Width
    gInfoC.HeightVal = shp.Height
    gInfoC.Rotation = shp.Rotation
    gInfoC.HFlip = shp.HorizontalFlip
    gInfoC.VFlip = shp.VerticalFlip
    ' fill
    gInfoC.HasFill = True
    gInfoC.FillVisible = shp.Fill.Visible
    If shp.Fill.Visible = msoTrue Then
        On Error Resume Next
        gInfoC.FillRGB = shp.Fill.ForeColor.RGB
        On Error GoTo 0
    End If
    ' frames
    gInfoC.HasLine = True
    gInfoC.LineVisible = shp.Line.Visible
    If shp.Line.Visible = msoTrue Then
        On Error Resume Next
        gInfoC.LineRGB = shp.Line.ForeColor.RGB
        gInfoC.LineWeight = shp.Line.Weight
        gInfoC.LineType = shp.Line.DashStyle
        On Error GoTo 0
    End If
    ' crop
    gInfoC.IsPicture = (shp.Type = msoPicture Or shp.Type = msoLinkedPicture)
    If gInfoC.IsPicture Then
        With shp.PictureFormat
            gInfoC.CropLeft = .CropLeft
            gInfoC.CropTop = .CropTop
            gInfoC.CropRight = .CropRight
            gInfoC.CropBottom = .CropBottom
        End With
    End If
    gInfoC.HasData = True
End Sub

Public Sub PasteShapeFormat()
    Dim sel As Selection
    Dim rng As ShapeRange
    Dim shp As Shape
    Dim targetA As Shape, targetB As Shape, targetC As Shape
    Dim n As Long
    'check how many formats copied
    If Not gInfoA.HasData And Not gInfoB.HasData And Not gInfoC.HasData Then
        MsgBox "Nothing copied yet", vbExclamation
        Exit Sub
    End If
    'check if target is shape
    Dim tmp As Shape
    Dim sld As Slide
    Set sel = ActiveWindow.Selection
    Set sld = ActiveWindow.View.Slide
    Select Case sel.Type
        Case ppSelectionShapes
            Set rng = sel.ShapeRange
        Case ppSelectionNone, ppSelectionSlides
            Set rng = sld.Shapes.Range
        Case Else
            MsgBox "select a shape to be applied", vbExclamation
            Exit Sub
    End Select
 
    For Each shp In rng
        If shp.Type <> msoPicture And shp.Type <> msoLinkedPicture Then
            GoTo NextIteration
        End If
        n = 0
        If gInfoA.HasData Then
            n = n + 1
            If n = 1 Then
                Set targetA = shp
            Else
                Set targetA = shp.Duplicate(1)
            End If
        End If
        If gInfoB.HasData Then
            n = n + 1
            If n = 1 Then
                Set targetB = shp
            Else
                Set targetB = shp.Duplicate(1)
            End If
        End If
        If gInfoC.HasData Then
            n = n + 1
            If n = 1 Then
                Set targetC = shp
            Else
                Set targetC = shp.Duplicate(1)
            End If
        End If
        If gInfoA.HasData Then ApplyPasteUnit targetA, gInfoA
        If gInfoB.HasData Then ApplyPasteUnit targetB, gInfoB
        If gInfoC.HasData Then ApplyPasteUnit targetC, gInfoC
NextIteration:
    Next shp
End Sub
 
Private Sub ApplyPasteUnit(ByVal shp As Shape, ByRef info As CopyShapeInfo)
    'unlock
    shp.LockAspectRatio = msoFalse
    'crop
    If info.IsPicture Then
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            With shp.PictureFormat
                .CropLeft = info.CropLeft
                .CropTop = info.CropTop
                .CropRight = info.CropRight
                .CropBottom = info.CropBottom
            End With
        End If
    End If
    'geometry
    shp.Left = info.Leftpos
    shp.Top = info.TopPos
    shp.Width = info.WidthVal
    shp.Height = info.HeightVal
    'rotate
    shp.Rotation = info.Rotation
    'flip
    On Error Resume Next
    If shp.HorizontalFlip <> info.HFlip Then
        shp.Flip msoFlipHorizontal
    End If
    If shp.VerticalFlip <> info.VFlip Then
        shp.Flip msoFlipVertical
    End If
    On Error GoTo 0
    'fill
    If info.HasFill Then
        shp.Fill.Visible = info.FillVisible
        If info.FillVisible = msoTrue Then
            On Error Resume Next
            shp.Fill.ForeColor.RGB = info.FillRGB
            On Error GoTo 0
        End If
    End If
    'frame
    If info.HasLine Then
        shp.Line.Visible = info.LineVisible
        If info.LineVisible = msoTrue Then
            On Error Resume Next
            shp.Line.ForeColor.RGB = info.LineRGB
            shp.Line.Weight = info.LineWeight
            shp.Line.DashStyle = info.LineType
            On Error GoTo 0
        End If
    End If
End Sub

Public Sub PasteII() 'to each slide first shape
    Dim startIndex As Long
    Dim sld As Slide
    Dim shp As Shape
    If Not gInfoA.HasData And Not gInfoB.HasData And Not gInfoC.HasData Then
        MsgBox "Nothing copied yet", vbExclamation
        Exit Sub
    End If
    startIndex = ActiveWindow.View.Slide.SlideIndex
    For Each sld In ActivePresentation.Slides
        If sld.SlideIndex > startIndex Then
            For Each shp In sld.Shapes
                If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                    sld.Select
                    shp.Select
                    Call PasteShapeFormat
                    Exit For
                End If
            Next shp
        End If
    Next sld
End Sub

Public Sub ClearInfo()
    gInfoA.HasData = False
    gInfoB.HasData = False
    gInfoC.HasData = False
End Sub

