' VERSION 1.0 CLASS
' BEGIN
'   MultiUse = -1  'True
' END
Option Explicit

'=============================================================================
' ArrayOps.cls — Generic array operations
'
' Depends on: VariantKit.cls (created locally)
'
' Public methods:
'   Sort, Slice, IndexOf, Flatten, SortIndices, CollectNumericColumns
'=============================================================================

' 延迟初始化的 VariantKit 单例（替代函数级 Static）
Private m_VK As VariantKit
Private Property Get VK() As VariantKit
    If m_VK Is Nothing Then
        Set m_VK = New VariantKit
    End If
    Set VK = m_VK
End Property

Public Sub Sort(ByRef arr As Variant, _
                Optional ByVal ascending As Boolean = True, _
                Optional ByVal comparer As String = "auto")
    If Not IsArray(arr) Then Exit Sub
    On Error Resume Next
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Sub
    Err.Clear
    ub = UBound(arr, 2)       ' 2D 检测 — Sort 仅支持 1D
    If Err.Number = 0 Then Err.Clear: On Error GoTo 0: Exit Sub
    Err.Clear: On Error GoTo 0
    If UBound(arr) <= LBound(arr) Then Exit Sub
    QuickSort arr, LBound(arr), UBound(arr), ascending, comparer
End Sub

Private Sub QuickSort(ByRef arr As Variant, ByVal low As Long, ByVal high As Long, _
                       ByVal ascending As Boolean, ByVal comparer As String)
    If high - low <= 16 Then
        InsertionSort arr, low, high, ascending, comparer: Exit Sub
    End If
    Dim i As Long, j As Long, pivot As Variant, tmp As Variant
    i = low: j = high: pivot = MedianOfThree(arr, low, high, comparer)
    Dim cmp As Long
    If ascending Then
        Do While i <= j
            Do While i <= high
                If CompareCmp(arr(i), pivot, comparer) >= 0 Then Exit Do
                i = i + 1
            Loop
            Do While j >= low
                If CompareCmp(pivot, arr(j), comparer) >= 0 Then Exit Do
                j = j - 1
            Loop
            If i <= j Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp: i = i + 1: j = j - 1
        Loop
    Else
        Do While i <= j
            Do While i <= high
                If CompareCmp(arr(i), pivot, comparer) <= 0 Then Exit Do
                i = i + 1
            Loop
            Do While j >= low
                If CompareCmp(pivot, arr(j), comparer) <= 0 Then Exit Do
                j = j - 1
            Loop
            If i <= j Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp: i = i + 1: j = j - 1
        Loop
    End If
    If (j - low) < (high - i) Then
        If low < j Then QuickSort arr, low, j, ascending, comparer
        If i < high Then QuickSort arr, i, high, ascending, comparer
    Else
        If i < high Then QuickSort arr, i, high, ascending, comparer
        If low < j Then QuickSort arr, low, j, ascending, comparer
    End If
End Sub

Private Sub InsertionSort(ByRef arr As Variant, ByVal low As Long, ByVal high As Long, _
                           ByVal ascending As Boolean, ByVal comparer As String)
    Dim i As Long, j As Long, key As Variant
    Dim sign As Long: If ascending Then sign = 1 Else sign = -1
    For i = low + 1 To high
        key = arr(i): j = i - 1
        Do While j >= low
            If sign * CompareCmp(arr(j), key, comparer) <= 0 Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = key
    Next i
End Sub

Private Function CompareCmp(ByVal a As Variant, ByVal b As Variant, _
                             ByVal comparer As String) As Long
    Select Case comparer
        Case "numeric"
            If IsNumeric(a) And IsNumeric(b) Then
                If CDbl(a) < CDbl(b) Then
                    CompareCmp = -1
                ElseIf CDbl(a) > CDbl(b) Then
                    CompareCmp = 1
                Else
                    CompareCmp = 0
                End If
            ElseIf IsNumeric(a) Then
                CompareCmp = 1
            ElseIf IsNumeric(b) Then
                CompareCmp = -1
            Else
                CompareCmp = StrComp(CStr(a), CStr(b), vbTextCompare)
            End If
        Case "text"
            CompareCmp = StrComp(CStr(a), CStr(b), vbTextCompare)
        Case Else
            ' VK 属性为延迟初始化单例，永不返回 Nothing — 直接委托
            ' (原 If vk Is Nothing 内联分支为死代码，已移除)
            ' ⚠ 行为变更: VK.Compare 对 Empty 和 Boolean 进行类型区分排序
            '   - Empty: 旧代码合并到数值零 (CDbl(Empty)=0)，新代码排在所有值之前
            '   - Boolean: 旧代码合并到数值 (CDbl(True)=-1)，新代码独立排序 (True>False, 排在非 Boolean 之后)
            '   新行为更正确，但属于排序语义变更 — 依赖旧排序的代码需评估
            CompareCmp = VK.Compare(a, b)
    End Select
End Function

' Median-of-three pivot selection — avoids O(n²) on sorted/patterned data
Private Sub ArrSwap(ByRef arr As Variant, ByVal i As Long, ByVal j As Long)
    Dim tmp As Variant: tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
End Sub

Private Sub IdxSwap(ByRef idx() As Long, ByVal i As Long, ByVal j As Long)
    Dim tmp As Long: tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
End Sub

Private Function MedianOfThree(ByRef arr As Variant, ByVal lo As Long, _
                                ByVal hi As Long, ByVal comparer As String) As Variant
    Dim mid As Long: mid = (lo + hi) \ 2
    If CompareCmp(arr(lo), arr(mid), comparer) > 0 Then ArrSwap arr, lo, mid
    If CompareCmp(arr(lo), arr(hi), comparer) > 0 Then ArrSwap arr, lo, hi
    If CompareCmp(arr(mid), arr(hi), comparer) > 0 Then ArrSwap arr, mid, hi
    MedianOfThree = arr(mid)
End Function

Private Function MedianOfThreeIdx(ByRef vals As Variant, ByRef idx() As Long, _
                                   ByVal lo As Long, ByVal hi As Long, _
                                   ByVal comp As String) As Variant
    Dim mid As Long: mid = (lo + hi) \ 2
    If CompareCmp(vals(idx(lo)), vals(idx(mid)), comp) > 0 Then IdxSwap idx, lo, mid
    If CompareCmp(vals(idx(lo)), vals(idx(hi)), comp) > 0 Then IdxSwap idx, lo, hi
    If CompareCmp(vals(idx(mid)), vals(idx(hi)), comp) > 0 Then IdxSwap idx, mid, hi
    MedianOfThreeIdx = vals(idx(mid))
End Function

Public Function Slice(ByRef arr As Variant, ByVal start As Long, _
                       Optional ByVal length As Long = -1) As Variant
    If Not IsArray(arr) Then Slice = Array(): Exit Function
    On Error Resume Next
    Dim lb As Long: lb = LBound(arr): Dim ub As Long: ub = UBound(arr)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Slice = Array(): Exit Function
    ub = UBound(arr, 2)       ' 2D 检测 — Slice 仅支持 1D
    If Err.Number = 0 Then Err.Clear: On Error GoTo 0: Slice = Array(): Exit Function
    Err.Clear: On Error GoTo 0
    Dim n As Long: n = ub - lb + 1
    ' Fix start before using it to compute length (avoid intermediate overflow)
    If start < 0 Then start = n + start
    If length < 0 Then length = n - start
    If start < 0 Or start >= n Or length <= 0 Then Slice = Array(): Exit Function
    If start + length > n Then length = n - start
    Dim result() As Variant, i As Long: ReDim result(0 To length - 1)
    For i = 0 To length - 1: result(i) = arr(lb + start + i): Next i
    Slice = result
End Function

Public Function IndexOf(ByRef arr As Variant, ByVal value As Variant) As Long
    IndexOf = -1
    If Not IsArray(arr) Then Exit Function
    Dim lb As Long, ub As Long, i As Long
    On Error Resume Next: lb = LBound(arr): ub = UBound(arr)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    ub = UBound(arr, 2)       ' 2D 检测 — IndexOf 仅支持 1D
    If Err.Number = 0 Then Err.Clear: On Error GoTo 0: Exit Function
    Err.Clear: On Error GoTo 0
    For i = lb To ub
        If VK.ValuesEqual(arr(i), value) Then IndexOf = i - lb: Exit Function
    Next i
End Function

Public Function Flatten(ByRef arr As Variant, _
                         Optional ByVal order As String = "R") As Variant
    Flatten = VK.Normalize1D(arr, order)
End Function

Public Sub SortIndices(ByRef values As Variant, ByRef indices() As Long, _
                        Optional ByVal asc As Boolean = True, _
                        Optional ByVal comparer As String = "auto")
    On Error Resume Next
    Dim lb As Long: lb = LBound(indices)
    Dim ub As Long: ub = UBound(indices)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Sub
    Err.Clear: On Error GoTo 0
    If ub <= lb Then Exit Sub
    QuickSortIdx values, indices, lb, ub, asc, comparer
End Sub

Private Sub QuickSortIdx(ByRef vals As Variant, ByRef idx() As Long, _
                          ByVal low As Long, ByVal high As Long, _
                          ByVal asc As Boolean, ByVal comp As String)
    If high - low <= 16 Then
        InsertionSortIdx vals, idx, low, high, asc, comp: Exit Sub
    End If
    Dim i As Long, j As Long, piv As Variant, tmp As Long
    i = low: j = high: piv = MedianOfThreeIdx(vals, idx, low, high, comp)
    If asc Then
        Do While i <= j
            Do While i <= high
                If CompareCmp(vals(idx(i)), piv, comp) >= 0 Then Exit Do
                i = i + 1
            Loop
            Do While j >= low
                If CompareCmp(piv, vals(idx(j)), comp) >= 0 Then Exit Do
                j = j - 1
            Loop
            If i <= j Then tmp = idx(i): idx(i) = idx(j): idx(j) = tmp: i = i + 1: j = j - 1
        Loop
    Else
        Do While i <= j
            Do While i <= high
                If CompareCmp(vals(idx(i)), piv, comp) <= 0 Then Exit Do
                i = i + 1
            Loop
            Do While j >= low
                If CompareCmp(piv, vals(idx(j)), comp) <= 0 Then Exit Do
                j = j - 1
            Loop
            If i <= j Then tmp = idx(i): idx(i) = idx(j): idx(j) = tmp: i = i + 1: j = j - 1
        Loop
    End If
    If (j - low) < (high - i) Then
        If low < j Then QuickSortIdx vals, idx, low, j, asc, comp
        If i < high Then QuickSortIdx vals, idx, i, high, asc, comp
    Else
        If i < high Then QuickSortIdx vals, idx, i, high, asc, comp
        If low < j Then QuickSortIdx vals, idx, low, j, asc, comp
    End If
End Sub

Private Sub InsertionSortIdx(ByRef vals As Variant, ByRef idx() As Long, _
                              ByVal low As Long, ByVal high As Long, _
                              ByVal asc As Boolean, ByVal comp As String)
    Dim i As Long, j As Long, key As Long
    Dim sign As Long: If asc Then sign = 1 Else sign = -1
    For i = low + 1 To high
        key = idx(i): j = i - 1
        Do While j >= low
            If sign * CompareCmp(vals(idx(j)), vals(key), comp) <= 0 Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = key
    Next i
End Sub

' CollectNumericColumns — identify columns where all non-empty data rows are numeric
' Returns: Long array of column indices (1-based). Empty array if no numeric columns.
Public Function CollectNumericColumns(ByRef dataArr As Variant, _
                                       ByVal numRows As Long, ByVal numCols As Long, _
                                       ByRef colNames() As String, _
                                       Optional ByVal hasHeader As Boolean = True, _
                                       Optional ByVal vk As VariantKit = Nothing) As Long()
    Dim c As Long, r As Long, v As Variant
    Dim isNum As Boolean
    Dim numColsFound As Long
    Dim colMap() As Long
    Dim firstDataRow As Long

    If vk Is Nothing Then Set vk = New VariantKit

    ReDim colNames(1 To numCols)
    numColsFound = 0
    If hasHeader Then firstDataRow = 2 Else firstDataRow = 1

    ReDim colMap(1 To numCols)  ' pre-allocate max, truncate after loop
    For c = 1 To numCols
        If hasHeader Then
            colNames(c) = CStr(dataArr(1, c))
        Else
            colNames(c) = "Col" & c
        End If
        isNum = True
        For r = firstDataRow To numRows
            v = dataArr(r, c)
            If Not IsEmpty(v) Then
                If Not vk.IsNumericCell(v) Then
                    isNum = False
                    Exit For
                End If
            End If
        Next
        If isNum Then
            numColsFound = numColsFound + 1
            colMap(numColsFound) = c
        End If
    Next

    If numColsFound = 0 Then
        Erase colMap
    ElseIf numColsFound < numCols Then
        ReDim Preserve colMap(1 To numColsFound)
    End If
    CollectNumericColumns = colMap
End Function

'=============================================================================
' Test_ArrayOps — Debug.Assert test harness
' Run: Call Test_ArrayOps (any assertion failure halts VBA)
'=============================================================================
Public Sub Test_ArrayOps()
    On Error GoTo ErrHandler
    Dim failCount As Long: failCount = 0
    Dim arr As Variant, result As Variant, arr2D As Variant
    Dim colMap() As Long, colNames() As String
    Dim i As Long

    '=====================================================================
    ' Sort — ascending
    '=====================================================================
    arr = Array(3, 1, 4, 1, 5)
    Sort arr, True
    If Not (arr(0) = 1) Then Err.Raise 5
    If Not (arr(1) = 1) Then Err.Raise 5
    If Not (arr(2) = 3) Then Err.Raise 5
    If Not (arr(3) = 4) Then Err.Raise 5
    If Not (arr(4) = 5) Then Err.Raise 5

    ' Sort — descending
    arr = Array(3, 1, 4, 1, 5)
    Sort arr, False
    If Not (arr(0) = 5) Then Err.Raise 5
    If Not (arr(1) = 4) Then Err.Raise 5
    If Not (arr(3) = 1) Then Err.Raise 5
    If Not (arr(4) = 1) Then Err.Raise 5

    ' Sort — already sorted (ascending)
    arr = Array(1, 2, 3, 4, 5)
    Sort arr, True
    If Not (arr(0) = 1) Then Err.Raise 5
    If Not (arr(4) = 5) Then Err.Raise 5

    ' Sort — reverse sorted
    arr = Array(9, 7, 5, 3, 1)
    Sort arr, True
    If Not (arr(0) = 1) Then Err.Raise 5
    If Not (arr(4) = 9) Then Err.Raise 5

    ' Sort — duplicates
    arr = Array(5, 5, 5, 1, 1)
    Sort arr, True
    If Not (arr(0) = 1) Then Err.Raise 5
    If Not (arr(1) = 1) Then Err.Raise 5
    If Not (arr(2) = 5) Then Err.Raise 5

    ' Sort — empty array (no crash — uninitialized array)
    Dim emptyArr() As Variant
    Sort emptyArr, True

    ' Sort — numeric comparer
    arr = Array("10", 2, "3")
    Sort arr, True, "numeric"
    If Not (arr(0) = 2) Then Err.Raise 5
    If Not (arr(1) = "3") Then Err.Raise 5
    If Not (arr(2) = "10") Then Err.Raise 5

    ' Sort — text comparer
    arr = Array("b", "a", "c")
    Sort arr, True, "text"
    If Not (arr(0) = "a") Then Err.Raise 5
    If Not (arr(1) = "b") Then Err.Raise 5
    If Not (arr(2) = "c") Then Err.Raise 5

    '=====================================================================
    ' Slice
    '=====================================================================
    arr = Array(10, 20, 30, 40, 50)

    ' Front slice
    result = Slice(arr, 0, 3)
    If Not (UBound(result) = 2) Then Err.Raise 5
    If Not (result(0) = 10) Then Err.Raise 5
    If Not (result(1) = 20) Then Err.Raise 5
    If Not (result(2) = 30) Then Err.Raise 5

    ' Middle slice
    result = Slice(arr, 1, 2)
    If Not (result(0) = 20) Then Err.Raise 5
    If Not (result(1) = 30) Then Err.Raise 5

    ' Back slice with negative start
    result = Slice(arr, -2)
    If Not (UBound(result) = 1) Then Err.Raise 5
    If Not (result(0) = 40) Then Err.Raise 5
    If Not (result(1) = 50) Then Err.Raise 5

    ' Negative start, explicit length
    result = Slice(arr, -1, 1)
    If Not (result(0) = 50) Then Err.Raise 5

    ' Length > available → clamped
    result = Slice(arr, 3, 10)
    If Not (UBound(result) = 1) Then Err.Raise 5
    If Not (result(0) = 40) Then Err.Raise 5
    If Not (result(1) = 50) Then Err.Raise 5

    ' Empty array → returns empty
    result = Slice(emptyArr, 0, 2)
    If Not (UBound(result) < LBound(result)) Then Err.Raise 5

    ' Start beyond bounds → empty
    result = Slice(arr, 100, 2)
    If Not (UBound(result) < LBound(result)) Then Err.Raise 5

    '=====================================================================
    ' IndexOf
    '=====================================================================
    arr = Array(10, 20, 30, 40, 50)
    If Not (IndexOf(arr, 10) = 0) Then Err.Raise 5
    If Not (IndexOf(arr, 30) = 2) Then Err.Raise 5
    If Not (IndexOf(arr, 50) = 4) Then Err.Raise 5
    If Not (IndexOf(arr, 99) = -1) Then Err.Raise 5

    ' Empty array → returns -1
    If Not (IndexOf(emptyArr, 1) = -1) Then Err.Raise 5

    '=====================================================================
    ' Flatten — 2D → 1D (row-major)
    '=====================================================================
    ReDim arr2D(1 To 2, 1 To 3)
    arr2D(1, 1) = 1: arr2D(1, 2) = 2: arr2D(1, 3) = 3
    arr2D(2, 1) = 4: arr2D(2, 2) = 5: arr2D(2, 3) = 6
    result = Flatten(arr2D)
    If Not (LBound(result) = 0) Then Err.Raise 5
    If Not (UBound(result) = 5) Then Err.Raise 5
    If Not (result(0) = 1) Then Err.Raise 5
    If Not (result(1) = 2) Then Err.Raise 5
    If Not (result(3) = 4) Then Err.Raise 5
    If Not (result(5) = 6) Then Err.Raise 5

    '=====================================================================
    ' CollectNumericColumns — identify all-numeric columns
    '=====================================================================
    ' Build a 4x3 array: Col1 all-numeric, Col2 mixed, Col3 all-string
    ReDim arr2D(1 To 4, 1 To 3)
    arr2D(1, 1) = "ID": arr2D(1, 2) = "Mixed": arr2D(1, 3) = "Label"
    arr2D(2, 1) = 10: arr2D(2, 2) = 5: arr2D(2, 3) = "A"
    arr2D(3, 1) = 20: arr2D(3, 2) = "X": arr2D(3, 3) = "B"
    arr2D(4, 1) = 30: arr2D(4, 2) = 7: arr2D(4, 3) = "C"

    colMap = CollectNumericColumns(arr2D, 4, 3, colNames, True)
    If Not (UBound(colMap) - LBound(colMap) + 1 = 1) Then Err.Raise 5
    If Not (colMap(1) = 1) Then Err.Raise 5
    If Not (colNames(1) = "ID") Then Err.Raise 5

    ' No header mode — use data without header strings
    Dim noHead(1 To 3, 1 To 2) As Variant
    noHead(1, 1) = 1:  noHead(1, 2) = "A"
    noHead(2, 1) = 2:  noHead(2, 2) = "B"
    noHead(3, 1) = 3:  noHead(3, 2) = "C"
    colMap = CollectNumericColumns(noHead, 3, 2, colNames, False)
    If Not (UBound(colMap) >= 1) Then Err.Raise 5
    If Not (colNames(1) = "Col1") Then Err.Raise 5

    ' All-string table → empty result
    Dim str2D(1 To 2, 1 To 1) As Variant
    str2D(1, 1) = "H": str2D(2, 1) = "X"
    colMap = CollectNumericColumns(str2D, 2, 1, colNames, True)
    ' Erase-d array → UBound errors = empty; catch and verify
    On Error Resume Next
    i = UBound(colMap)
    If Err.Number = 0 Then Err.Raise 5  ' should error if truly empty
    Err.Clear: On Error GoTo 0

    '=====================================================================
    ' Median-of-three pivot — sorted/reverse-sorted (regression: fix C4)
    '=====================================================================
    ' Already sorted — median-of-three prevents O(n^2)
    arr = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
    Sort arr, True
    If Not (arr(0) = 1) Then Err.Raise 5
    If Not (arr(19) = 20) Then Err.Raise 5

    ' Reverse sorted
    arr = Array(20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1)
    Sort arr, True
    If Not (arr(0) = 1) Then Err.Raise 5
    If Not (arr(19) = 20) Then Err.Raise 5

    ' All same — should not infinite-loop
    arr = Array(5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5)
    Sort arr, True
    If Not (arr(0) = 5) Then Err.Raise 5
    If Not (arr(19) = 5) Then Err.Raise 5

    ' CollectNumericColumns — all numeric (regression: fix C3)
    ReDim allNum(1 To 3, 1 To 3)
    allNum(1, 1) = "H1": allNum(1, 2) = "H2": allNum(1, 3) = "H3"
    allNum(2, 1) = 1:     allNum(2, 2) = 2:     allNum(2, 3) = 3
    allNum(3, 1) = 4:     allNum(3, 2) = 5:     allNum(3, 3) = 6
    colMap = CollectNumericColumns(allNum, 3, 3, colNames, True)
    If Not (UBound(colMap) - LBound(colMap) + 1 = 3) Then Err.Raise 5
    If Not (colMap(1) = 1) Then Err.Raise 5
    If Not (colMap(2) = 2) Then Err.Raise 5
    If Not (colMap(3) = 3) Then Err.Raise 5

    If failCount > 0 Then Err.Raise 5, , "ArrayOps: " & failCount & " assertion(s) failed"
    Exit Sub
ErrHandler:
    If Err.Number = 5 Then
        failCount = failCount + 1
        Err.Clear
        Resume Next
    End If
    Err.Raise Err.Number
End Sub