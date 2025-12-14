Option Explicit
 
Public Type Point2D
    x As Double
    y As Double
End Type
Public Sub FineGriding()
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
    Debug.Print "bestrows:"; bestrows; "bestcols"; bestcols; "bestscore:"; bestscore
    '3. Sort the points by y,x'
    Dim order() As Long
    ReDim order(1 To n)
    OrderByRowsCols pts, order, bestrows, bestcols, n
    '4. Find avg spacing dx and dy'
    Dim dx As Double, dy As Double
    GroupSpacing pts, order, bestrows, bestcols, dx, dy
    '5. Apply griding, anchoring the left top shape, use avg dy dx as separations'
    Dim x_pos As Double, y_pos As Double
    Dim idx As Long
    For i = 1 To bestcols
        x_pos = pts(order(1)).x + (i - 1) * dx
        For j = 1 To bestrows
            y_pos = pts(order(1)).y + (j - 1) * dy
            idx = (i - 1) * bestrows + j
            With sr(order(idx))
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
    Dim stdy As Double, stdx As Double
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
    Groupstd pts, rows, cols, stdx, stdy
    'Debug.Print "   Score: stdy="; stdy; " rows = "; rows; "stdx = "; stdx; "cols = "; cols
    GridScoreForShape = stdy + stdx
End Function
 
Public Sub Groupstd(ByRef pts() As Point2D, ByVal rows As Long, ByVal cols As Long, ByRef avgstdx As Double, ByRef avgstdy As Double)
    Dim i As Long, j As Long, n As Long, index As Long
    Dim test_ord() As Long
    Dim list() As Double
 
    n = UBound(pts) - LBound(pts) + 1
    ReDim test_ord(1 To n)
    OrderByRowsCols pts, test_ord, rows, cols, n
    ' works on group stdx
    avgstdx = 0
    For i = 1 To cols
        Erase list
        ReDim list(1 To rows)
        For j = 1 To rows
            index = (i - 1) * rows + j
            list(j) = pts(test_ord(index)).x
        Next j
        avgstdx = avgstdx + ComputeStd(list) / cols
    Next i
    ' works on group stdy
    avgstdy = 0
    For i = 1 To rows
        Erase list
        ReDim list(1 To cols)
        For j = 1 To cols
            index = (j - 1) * rows + i
            list(j) = pts(test_ord(index)).y
        Next j
        avgstdy = avgstdy + ComputeStd(list) / rows
    Next i
End Sub
Public Sub GroupSpacing(ByRef pts() As Point2D, ByRef order() As Long, ByVal rows As Long, ByVal cols As Long, ByRef dx As Double, ByRef dy As Double)
    Dim i As Long, j As Long
    Dim startidx As Long, endidx As Long
    Dim index1 As Long, index2 As Long
    Dim tempx As Double, tempy As Double
    dx = 0
    dy = 0
    ' works on dy, y separation
    For i = 1 To cols
        startidx = (i - 1) * rows + 1
        endidx = i * rows
        tempy = 0
        For j = startidx To endidx - 1
            tempy = tempy + (pts(order(j + 1)).y - pts(order(j)).y) / (rows - 1)
        Next j
        dy = dy + tempy / cols
    Next i
    ' works on dx, x separation
    For i = 1 To rows
        tempx = 0
        For j = 1 To cols - 1
            index1 = (j - 1) * rows + i
            index2 = j * rows + i
            tempx = tempx + (pts(order(index2)).x - pts(order(index1)).x) / (cols - 1)
        Next j
        dx = dx + tempx / rows
    Next i
End Sub
Public Sub OrderByRowsCols(ByRef pts() As Point2D, ByRef order() As Long, ByVal rows As Long, ByVal cols As Long, ByVal n As Long)
    Dim startidx As Long, endidx As Long
    Dim xs() As Double, ys() As Double
    Dim i As Long
    ReDim xs(1 To n)
    ReDim ys(1 To n)
    For i = 1 To n
        xs(i) = pts(i).x
        ys(i) = pts(i).y
        order(i) = i
    Next i
    SortArrayIndex order, xs, 1, n
    For i = 1 To cols
        startidx = (i - 1) * rows + 1
        endidx = i * rows
        SortArrayIndex order, ys, startidx, endidx
    Next i
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
Public Sub ComputeMeanStd(ByRef arr() As Double, ByVal startidx As Long, ByVal endidx As Long, ByRef stdval As Double)
    Dim i As Long, n As Long
    Dim s As Double
    Dim diff As Double
    Dim meanval As Double, varsum As Double
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
    varsum = 0
    For i = startidx To endidx
        diff = arr(i) - meanval
        varsum = varsum + diff * diff
    Next i
    stdval = Sqr(varsum / n)
End Sub
Public Function ComputeStd(ByRef arr() As Double) As Double
    Dim i As Long, n As Long
    Dim diff As Double
    Dim meanval As Double, varsum As Double
    n = UBound(arr) - LBound(arr) + 1
    meanval = 0
    For i = 1 To n
        meanval = meanval + arr(i) / n
    Next i
    varsum = 0
    For i = 1 To n
        diff = arr(i) - meanval
        varsum = varsum + diff * diff
    Next i
    ComputeStd = Sqr(varsum / n)
End Function
