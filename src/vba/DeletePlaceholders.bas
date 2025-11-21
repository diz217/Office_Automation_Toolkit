Option Explicit

Public Sub Deleteplaceholdertitle()
    Dim startIdx As Long
    Dim i As Long
    Dim sld As Slide
    Dim shp As Shape
    Dim txt As String
    startIdx = ActiveWindow.View.Slide.SlideIndex
    For i = startIdx To ActivePresentation.Slides.count
        Set sld = ActivePresentation.Slides(i)
        For Each shp In sld.Shapes
            If shp.Type = msoPlaceholder Then
                Select Case shp.PlaceholderFormat.Type
                    Case ppPlaceholderTitle, ppPlaceholderCenterTitle
                        On Error Resume Next
                        txt = ""
                        If shp.HasTextFrame Then
                            If shp.TextFrame.HasText Then
                                txt = Trim(shp.TextFrame.TextRange.Text)
                            End If
                        End If
                        On Error GoTo 0
                        If txt = "" Then
                            shp.Delete
                        End If
                End Select
            End If
        Next shp
    Next i
End Sub
