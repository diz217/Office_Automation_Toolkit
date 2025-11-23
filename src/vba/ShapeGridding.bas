Option Explicit

Public Type Point2D
    x As Double
    y As Double
End Type
Public Sub FineGridding()
    Dim sel As Selection
    Dim sr As ShapeRange
    Dim n As Long
    Dim pts() As Point2D
    Dim i As Long, j As Long
    Dim bestrows As Long, bestcols As Long
    Dim bestscore As Double
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        MsgBox "Select some shapes", vbExclamation
        Exit Sub
    End If
    Set sr = sel.ShapeRange
    n = sr.count
    If n = 0 Then
         MsgBox "None shapes selected", vbExclamation
         Exit Sub
    ElseIf n = 1 Then
        Exit Sub
    End If
    '1. Initialize point2D set'
    ReDim pts(1 To n)
    For i = 1 To n
        With sr(i)
            pts(i).x = .Left
            pts(i).y = .Top
        End With
    Next i
    '2. Find best grid: bestrows, bestcols '
    FindBestGrid pts, bestrows, bestcols, bestscore
    '3. Find avg spacing dx and dy'
    Dim xs() As Double, ys() As Double
    Dim order() As Long
    Dim stdx As Double, stdy As Double
    Dim dx As Double, dy As Double
    ReDim xs(1 To n)
    ReDim ys(1 To n)
    ReDim order(1 To n)
    For i = 1 To n
        xs(i) = pts(i).x
        ys(i) = pts(i).y
        order(i) = i
    Next i
    GroupStdAndSpacing ys, bestrows, stdy, dy
    GroupStdAndSpacing xs, bestcols, stdx, dx
    '4. Sort the points by y,x'
    Dim startidx As Long, endidx As Long
    SortArrayIndex order, xs, 1, n
    For i = 1 To bestcols
        startidx = (i - 1) * bestrows + 1
        endidx = i * bestrows
        SortArrayIndex order, ys, startidx, endidx
    Next i
    '5. Apply griding, anchoring the left top shape, use avg dy dx as separations'
    Dim x_pos As Double, y_pos As Double
    Dim idx As Long
    For i = 1 To bestcols
        x_pos = pts(order(1)).x + (i - 1) * dx
        For j = 1 To bestrows
            y_pos = pts(order(1)).y + (j - 1) * dy
            idx = (i - 1) * bestrows + j
            With sr(idx)
                .Left = x_pos
                .Top = y_pos
            End With
        Next j
    Next i
End Sub
 
Public Sub InferGridFromSelection()
    Dim sel As Selection
    Dim sr As ShapeRange
    Dim n As Long
    Dim pts() As Point2D
    Dim i As Long
    Dim bestrows As Long, bestcols As Long
    Dim bestscore As Double
    Set sel = ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes Then
        MsgBox "Select some shapes", vbExclamation
        Exit Sub
    End If
    Set sr = sel.ShapeRange
    n = sr.count
    If n = 0 Then
         MsgBox "None shapes selected", vbExclamation
         Exit Sub
    End If
    ReDim pts(1 To n)
    For i = 1 To n
        With sr(i)
            pts(i).x = .Left + .Width / 2
            pts(i).y = .Top + .Height / 2
        End With
    Next i
    FindBestGrid pts, bestrows, bestcols, bestscore
    MsgBox "Best grid: " & bestrows & "x" & bestcols & ", score is " & bestscore
End Sub
 
Public Sub FindBestGrid(ByRef pts() As Point2D, ByRef bestrows As Long, ByRef bestcols As Long, ByRef bestscore As Double)
    Dim n As Long
    Dim gridrows() As Long, gridcols() As Long
    Dim numgrids As Long
    Dim i As Long
    Dim j As Long
    Dim r As Long, c As Long
    Dim score As Double
    n = UBound(pts) - LBound(pts) + 1
    BuildGridCandidates n, gridrows, gridcols, numgrids
    bestscore = 10000000#
    bestrows = 1
    bestcols = n
    For i = 1 To numgrids
        r = gridrows(i)
        c = gridcols(i)
        score = GridScoreForShape(pts, r, c)
        Debug.Print "score="; score; "; bestscore="; bestscore
        If score < bestscore Then
            bestscore = score
            bestrows = r
            bestcols = c
        End If
    Next i
End Sub

Public Sub BuildGridCandidates(ByVal n As Long, ByRef gridrows() As Long, ByRef gridcols() As Long, ByRef numgrids As Long)
    Dim r As Long
    numgrids = 0
    ReDim gridrows(1 To 1)
    ReDim gridcols(1 To 1)
    For r = 1 To n
        If n Mod r = 0 Then
            numgrids = numgrids + 1
            ReDim Preserve gridrows(1 To numgrids)
            ReDim Preserve gridcols(1 To numgrids)
            gridrows(numgrids) = r
            gridcols(numgrids) = n \ r
        End If
    Next r
End Sub
 
Public Function GridScoreForShape(ByRef pts() As Point2D, ByVal rows As Long, ByVal cols As Long) As Double
    Dim n As Long
    Dim xs() As Double, ys() As Double
    Dim i As Long
    Dim stdy As Double, dy As Double
    Dim stdx As Double, dx As Double
    Dim eps As Double
    n = UBound(pts) - LBound(pts) + 1
    If rows * cols <> n Then
        GridScoreForShape = 10000000#
        Exit Function
    End If
    ReDim xs(1 To n)
    ReDim ys(1 To n)
    For i = 1 To n
        xs(i) = pts(i).x
        ys(i) = pts(i).y
    Next i
    GroupStdAndSpacing ys, rows, stdy, dy
    GroupStdAndSpacing xs, cols, stdx, dx
    Debug.Print "stdy="; stdy; "; dy="; dy; "rows="; rows; "stdx="; stdx; "; dx="; dx; "cols="; cols
    eps = 0.000001
    GridScoreForShape = stdy / (dy + eps) + stdx / (dx + eps)
End Function
 
Public Sub GroupStdAndSpacing(ByRef values() As Double, ByVal groups As Long, ByRef avgstd As Double, ByRef avgspacing As Double)
    Dim n As Long
    Dim sorted() As Double
    Dim i As Long
    Dim base As Long
    Dim startidx As Long, endidx As Long, g As Long
    Dim chunkmean As Double, chunkstd As Double
    Dim groupmeans() As Double, groupstds() As Double
    Dim diffs() As Double
    n = UBound(values) - LBound(values) + 1
    If groups <= 0 Or groups > n Then
        avgstd = 0
        avgspacing = 1
    End If
    ReDim sorted(1 To n)
    For i = 1 To n
        sorted(i) = values(i)
    Next i
    SortArray sorted
    base = n \ groups
    ReDim groupmeans(1 To groups)
    ReDim groupstds(1 To groups)
    startidx = 1
    For g = 1 To groups
        endidx = startidx + base - 1
        ComputeMeanStd sorted, startidx, endidx, chunkmean, chunkstd
        groupmeans(g) = chunkmean
        groupstds(g) = chunkstd
        startidx = endidx + 1
    Next g
    avgstd = 0
    For g = 1 To groups
        avgstd = avgstd + groupstds(g)
    Next g
    avgstd = avgstd / groups
    If groups = 1 Then
        avgspacing = sorted(n) - sorted(1)
        If avgspacing <= 0 Then avgspacing = 0.000001
    Else
        SortArray groupmeans
        ReDim diffs(1 To groups - 1)
        For g = 1 To groups - 1
            diffs(g) = groupmeans(g + 1) - groupmeans(g)
        Next g
        avgspacing = 0
        For g = 1 To groups - 1
            avgspacing = avgspacing + diffs(g)
        Next g
        avgspacing = avgspacing / (groups - 1)
        If avgspacing <= 0 Then avgspacing = 0.000001
    End If
End Sub
 
Public Sub SortArray(ByRef arr() As Double)
    Dim i As Long, j As Long
    Dim tmp As Double
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(j) < arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub
Public Sub SortArrayIndex(ByRef ord() As Long, ByRef arr() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim tmp As Long
 
    For i = lo To hi - 1
        For j = i + 1 To hi
            If arr(ord(j)) < arr(ord(i)) Then
                tmp = ord(i)
                ord(i) = ord(j)
                ord(j) = tmp
            End If
        Next j
    Next i
End Sub
Public Sub ComputeMeanStd(ByRef arr() As Double, ByVal startidx As Long, ByVal endidx As Long, ByRef meanval As Double, ByRef stdval As Double)
    Dim i As Long, n As Long
    Dim s As Double
    Dim diff As Double
    Dim varsum As Double
    n = endidx - startidx + 1
    If n <= 0 Then
        meanval = 0
        stdval = 0
        Exit Sub
    End If
    s = 0
    For i = startidx To endidx
        s = s + arr(i)
    Next i
    meanval = s / n
    If n = 1 Then
        stdval = 0
    End If
    varsum = 0
    For i = startidx To endidx
        diff = arr(i) - meanval
        varsum = varsum + diff * diff
    Next i
    stdval = Sqr(varsum / n)
End Sub
