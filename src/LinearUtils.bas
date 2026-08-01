Option Explicit

'==============================================================================
' Module:       LinearUtils
' Purpose:      Linear algebra: SVD, QR, LU, Cholesky, PINV, eigenvalues
' Layer:        Statistics
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       50 functions/subs
'==============================================================================


'=====================================================================
' LinearUtils.bas — VBA 线性代数标准库
'
' 约定: 所有输出矩阵均为 1-based (与 Excel Range 一致)。
' 调用者注意: ArrayUtils 模块输出为 0-based, 传入矩阵函数前需转换基址。
'
' 基础:
'   MatrixRows, MatrixCols, MatrixFrobeniusNorm, MatrixTranspose,
'   IdentityMatrix, MatrixCopy, MatrixGetColumn, MatrixSetColumn,
'   MatrixTrace, MatrixScale, MatrixAdd, MatrixSubtract, MatrixHadamard,
'   MatrixNorm, MatrixPower
' 线性代数:
'   MatrixMultiply, MatrixMultiplyNaive
'   LUDecomposition, MatrixDeterminant, SolveLinearSystem, MatrixConditionNumber
' 分解:
'   SVD, PseudoInverse, MatrixRank_Array, EigenSymmetric, QRDecomposition,
'   LUDecomposition, CholeskyDecomposition
' 向量:
'   VectorDot, VectorNorm, VectorCross
' 转换:
'   RangeToMatrix, MatrixToRange, SelectionToArray2D, ArrayToRange
'
' 工作表 UDF (可直接在单元格中使用):
'   UDF_LINALG_SVD_U, UDF_LINALG_SVD_S, UDF_LINALG_SVD_VT, UDF_LINALG_SVD_SVALS,
'   UDF_LINALG_PINV, UDF_LINALG_QR_Q, UDF_LINALG_QR_R,
'   UDF_LINALG_DET, UDF_LINALG_SOLVE, UDF_LINALG_CHOLESKY,
'   UDF_LINALG_EIGVAL, UDF_LINALG_EIGVEC, UDF_LINALG_RANK, UDF_LINALG_POLYFIT
'=====================================================================

'=============================================================================
' 常量
'=============================================================================
Private Const NUM_EPSILON As Double = 2.22044604925031E-16
Private Const DEFAULT_TOL   As Double = 0.00000000000001
Private Const MAX_SWEEPS    As Long = 50
Private Const MAX_DOUBLE    As Double = 1.79769313486231E+308
Private Const ERR_INVALID_INPUT  As Long = vbObjectError + 1001
Private Const ERR_DIM_MISMATCH   As Long = vbObjectError + 1002
Private Const ERR_MULTI_AREA     As Long = vbObjectError + 1003
Private Const ERR_INVALID_SIZE   As Long = vbObjectError + 2000
Private Const ERR_NAN_INF        As Long = vbObjectError + 2001
Private Const ERR_SWEEP_LIMIT     As Long = vbObjectError + 2003
Private Const ERR_SVD_CONVERGE   As Long = vbObjectError + 2004
Private Const ERR_EIGEN_CONVERGE As Long = vbObjectError + 2005
Private Const ERR_RANGE_NOTHING   As Long = vbObjectError + 1000
Private Const ERR_DIM_INVALID     As Long = vbObjectError + 1005
Private Const ERR_NOT_SYMMETRIC   As Long = vbObjectError + 1006
Private Const ERR_SINGULAR        As Long = vbObjectError + 1061
Private Const ERR_SOLVE_MISMATCH  As Long = vbObjectError + 1070
Private Const ERR_VEC_DIM         As Long = vbObjectError + 1080
Private Const ERR_VEC_3D_ONLY     As Long = vbObjectError + 1081
Private Const ERR_CHOLESKY_SIZE   As Long = vbObjectError + 1100
Private Const ERR_NOT_POS_DEF     As Long = vbObjectError + 1101
Private Const ERR_CHOLESKY_NOT_SYM As Long = vbObjectError + 1102
Private Const ERR_MATRIX_TOO_BIG  As Long = vbObjectError + 2006
Private Const ERR_COL_OUT_OF_RANGE As Long = vbObjectError + 2007
Private Const ERR_COL_LEN_MISMATCH As Long = vbObjectError + 2008
Private Const ERR_SETCOL_LEN      As Long = vbObjectError + 2009
Private Const ERR_IDENTITY_SIZE   As Long = vbObjectError + 2010
Private Const ERR_NOT_SQUARE      As Long = vbObjectError + 2100
Private Const ERR_ADD_MISMATCH    As Long = vbObjectError + 2101
Private Const ERR_SUB_MISMATCH    As Long = vbObjectError + 2102
Private Const ERR_HADAMARD_MISMATCH As Long = vbObjectError + 2103
Private Const ERR_POWER_SQUARE    As Long = vbObjectError + 2104
Private Const ERR_NEG_EXPONENT    As Long = vbObjectError + 2105
Private Const ERR_LU_SQUARE       As Long = vbObjectError + 1060

'=============================================================================
' 私有辅助函数
'=============================================================================
Private Function IsFiniteDouble(ByVal x As Double) As Boolean
    IsFiniteDouble = (x = x) And (Abs(x) <= MAX_DOUBLE)
End Function

Private Function MinLong(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinLong = a Else MinLong = b
End Function

' 返回 sign(x); x>=0 返回 +1, x<0 返回 -1 (零视为正, Jacobi 旋转可接受任意符号)
Private Function SignNonZero(ByVal x As Double) As Double
    If x >= 0# Then SignNonZero = 1# Else SignNonZero = -1#
End Function

Private Sub ValidateMatrix(ByRef A As Variant, ByVal matrixName As String)
    ' Extract data to local — do NOT mutate ByRef A (§1.1)
    Dim matData As Variant
    If IsObject(A) Then
        If TypeOf A Is Range Then matData = A.Value Else matData = A
    Else
        matData = A
    End If
    Dim m As Long, n As Long, i As Long, j As Long
    m = UBound(matData, 1) - LBound(matData, 1) + 1
    n = UBound(matData, 2) - LBound(matData, 2) + 1
    If m < 1 Or n < 1 Then
        Err.Raise ERR_INVALID_SIZE, "ValidateMatrix", _
            "无效矩阵尺寸: " & matrixName & " 为 " & m & "x" & n
    End If
    Dim r0 As Long: r0 = LBound(matData, 1)
    Dim c0 As Long: c0 = LBound(matData, 2)
    Dim cellVal As Variant
    For i = 1 To m
        For j = 1 To n
            cellVal = matData(r0 + i - 1, c0 + j - 1)
            If IsError(cellVal) Then
                Err.Raise ERR_INVALID_INPUT, "ValidateMatrix", _
                    "矩阵 " & matrixName & " 包含 Error 值于 (" & i & "," & j & ")"
            End If
            If Not IsNumeric(cellVal) Then
                Err.Raise ERR_INVALID_INPUT, "ValidateMatrix", _
                    "矩阵 " & matrixName & " 包含非数值于 (" & i & "," & j & ")"
            End If
            If Not IsFiniteDouble(CDbl(cellVal)) Then
                Err.Raise ERR_NAN_INF, "ValidateMatrix", _
                    "矩阵 " & matrixName & " 包含 NaN/Inf 于 (" & i & "," & j & ")"
            End If
        Next j
    Next i
End Sub

'=============================================================================
' 内部辅助 — Variant → Double() 转换 (COM 兼容)
'=============================================================================
Private Function ToDoubleMatrix(ByRef v As Variant) As Double()
    Dim result() As Double, i As Long, j As Long
    Dim lb1 As Long, lb2 As Long, ub1 As Long, ub2 As Long

    ' Extract to local — do NOT mutate ByRef v (§1.1)
    Dim localV As Variant
    If IsObject(v) Then
        If TypeOf v Is Range Then localV = v.Value Else localV = v
    Else
        localV = v
    End If

    If Not IsArray(localV) Then
        ReDim result(1 To 1, 1 To 1): result(1, 1) = CDbl(localV)
        ToDoubleMatrix = result: Exit Function
    End If
    lb1 = LBound(localV, 1): ub1 = UBound(localV, 1)
    lb2 = LBound(localV, 2): ub2 = UBound(localV, 2)
    ReDim result(1 To ub1 - lb1 + 1, 1 To ub2 - lb2 + 1)
    For i = 1 To UBound(result, 1)
        For j = 1 To UBound(result, 2)
            result(i, j) = CDbl(localV(lb1 + i - 1, lb2 + j - 1))
        Next j
    Next i
    ToDoubleMatrix = result
End Function

'=============================================================================
' 公共基础函数
'=============================================================================
Public Function MatrixRows(ByRef A As Variant) As Long
    ' 提取局部变量 — 不修改 ByRef 参数 (§1.1)
    Dim localA As Variant
    If IsObject(A) Then
        If TypeOf A Is Range Then localA = A.Value Else localA = A
    Else
        localA = A
    End If
    ' 快速路径: 已是二维数组则直接用 UBound (避免 ToDoubleMatrix 深拷贝)
    If IsArray(localA) Then
        On Error Resume Next
        Dim ub2 As Long: ub2 = UBound(localA, 2)
        If Err.Number = 0 Then
            MatrixRows = UBound(localA, 1) - LBound(localA, 1) + 1
            On Error GoTo 0: Exit Function
        End If
        Err.Clear: On Error GoTo 0
        ' 1D 数组 — 直接返回元素数 (避免 ToDoubleMatrix 在 LBound(,2) 上崩溃)
        MatrixRows = UBound(localA, 1) - LBound(localA, 1) + 1
        Exit Function
    End If
    Dim mat() As Double: mat = ToDoubleMatrix(localA)
    MatrixRows = UBound(mat, 1)
End Function

Public Function MatrixCols(ByRef A As Variant) As Long
    ' 提取局部变量 — 不修改 ByRef 参数 (§1.1)
    Dim localA As Variant
    If IsObject(A) Then
        If TypeOf A Is Range Then localA = A.Value Else localA = A
    Else
        localA = A
    End If
    ' 快速路径: 已是二维数组则直接用 UBound (避免 ToDoubleMatrix 深拷贝)
    If IsArray(localA) Then
        On Error Resume Next
        Dim ub2c As Long: ub2c = UBound(localA, 2)
        If Err.Number = 0 Then
            MatrixCols = UBound(localA, 2) - LBound(localA, 2) + 1
            On Error GoTo 0: Exit Function
        End If
        Err.Clear: On Error GoTo 0
        ' 1D 数组 — 返回 1 (单列矩阵)
        MatrixCols = 1
        Exit Function
    End If
    Dim mat() As Double: mat = ToDoubleMatrix(localA)
    MatrixCols = UBound(mat, 2)
End Function

Public Function MatrixFrobeniusNorm(ByRef A As Variant) As Double
    Dim matA() As Double: matA = ToDoubleMatrix(A)
    ValidateMatrix matA, "A"
    Dim m As Long, n As Long, i As Long, j As Long
    m = MatrixRows(matA): n = MatrixCols(matA)

    ' 找出最大绝对值用于缩放, 防止平方运算溢出
    Dim maxAbs As Double: maxAbs = 0#
    For i = 1 To m
        For j = 1 To n
            If Abs(matA(i, j)) > maxAbs Then maxAbs = Abs(matA(i, j))
        Next j
    Next i
    If maxAbs = 0# Then
        MatrixFrobeniusNorm = 0#
        Exit Function
    End If

    Dim s As Double: s = 0#
    For i = 1 To m
        For j = 1 To n
            s = s + (matA(i, j) / maxAbs) ^ 2
        Next j
    Next i
    MatrixFrobeniusNorm = maxAbs * Sqr(s)
End Function

'=============================================================================
' Range 与 Matrix 互转
'=============================================================================
Private Function IsNumericCell(ByVal v As Variant) As Boolean
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    IsNumericCell = vk.IsNumericCell(v)
End Function

Public Function RangeToMatrix(ByVal rng As Variant) As Double()
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    If Not TypeOf rng Is Range Then
        If IsArray(rng) Then
            Dim tmpDoubles() As Double
            tmpDoubles = vk.ToDoubles(rng)
            ' Ensure always 2D (matrix convention)
            Dim dLb As Long, dUb As Long
            On Error Resume Next
            dLb = LBound(tmpDoubles, 2)
            If Err.Number <> 0 Then
                ' 1D → wrap as single-row 2D matrix (1 To 1, 1 To n)
                Err.Clear: On Error GoTo 0
                Dim rowMat() As Double
                Dim dLen As Long: dLen = UBound(tmpDoubles) - LBound(tmpDoubles) + 1
                ReDim rowMat(1 To 1, 1 To dLen)
                Dim di As Long
                For di = 1 To dLen
                    rowMat(1, di) = tmpDoubles(LBound(tmpDoubles) + di - 1)
                Next di
                RangeToMatrix = rowMat
            Else
                On Error GoTo 0
                RangeToMatrix = tmpDoubles
            End If
            Exit Function
        End If
        Err.Raise ERR_RANGE_NOTHING, "RangeToMatrix", "输入不是 Range 或数组。"
    End If
    If rng Is Nothing Then
        Err.Raise ERR_RANGE_NOTHING, "RangeToMatrix", "输入区域为 Nothing。"
    End If
    If rng.Areas.Count > 1 Then
        Err.Raise ERR_MULTI_AREA, "RangeToMatrix", "不支持多选区，请使用连续区域。"
    End If
    Dim result() As Double
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim v As Variant, cellVal As Variant
    nRows = rng.Rows.Count
    nCols = rng.Columns.Count
    ReDim result(1 To nRows, 1 To nCols)
    If nRows = 1 And nCols = 1 Then
        cellVal = rng.Value
        If Not IsNumericCell(cellVal) Then
            Err.Raise ERR_INVALID_INPUT, "RangeToMatrix", _
                "单元格 " & rng.Address(False, False) & " 包含非数值数据。"
        End If
        result(1, 1) = CDbl(cellVal)
    Else
        v = rng.Value
        Dim baseRow As Long: baseRow = rng.Row
        Dim baseCol As Long: baseCol = rng.Column
        Dim ws As Worksheet: Set ws = rng.Parent
        For i = 1 To nRows
            For j = 1 To nCols
                cellVal = v(i, j)
                If Not IsNumericCell(cellVal) Then
                    Err.Raise ERR_INVALID_INPUT, "RangeToMatrix", _
                        "单元格 " & ws.Cells(baseRow + i - 1, baseCol + j - 1).Address(False, False) & " 包含非数值数据。"
                End If
                result(i, j) = CDbl(cellVal)
            Next j
        Next i
    End If
    RangeToMatrix = result
End Function

Public Sub MatrixToRange(ByRef mat() As Double, ByRef startCell As Range)
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim outArr() As Double
    Dim r0 As Long, c0 As Long
    nRows = MatrixRows(mat)
    nCols = MatrixCols(mat)
    If LBound(mat, 1) = 1 And LBound(mat, 2) = 1 Then
        startCell.Resize(nRows, nCols).Value = mat
    Else
        ReDim outArr(1 To nRows, 1 To nCols)
        r0 = LBound(mat, 1)
        c0 = LBound(mat, 2)
        For i = 1 To nRows
            For j = 1 To nCols
                outArr(i, j) = mat(r0 + i - 1, c0 + j - 1)
            Next j
        Next i
        startCell.Resize(nRows, nCols).Value = outArr
    End If
End Sub

Private Function DblMatrixToVariant(ByRef mat() As Double) As Variant
    Dim m As Long, n As Long, i As Long, j As Long
    m = MatrixRows(mat): n = MatrixCols(mat)
    Dim result() As Variant
    ReDim result(1 To m, 1 To n)
    Dim r0 As Long: r0 = LBound(mat, 1)
    Dim c0 As Long: c0 = LBound(mat, 2)
    For i = 1 To m
        For j = 1 To n
            result(i, j) = mat(r0 + i - 1, c0 + j - 1)
        Next j
    Next i
    DblMatrixToVariant = result
End Function

'=============================================================================
' Range 与 Variant 数组互转 (通用版本，保留所有数据类型)
'=============================================================================

' 将选中单元格区域转换为 2D Variant 数组。
' 适配: 单个单元格、单行、单列、多行多列。
' 结果始终为 1-based 2D 数组: (1 To 行数, 1 To 列数)。
Public Function SelectionToArray2D(ByRef rng As Range) As Variant()
    Dim result() As Variant

    If rng Is Nothing Then
        ReDim result(1 To 1, 1 To 1): Erase result
        SelectionToArray2D = result
        Exit Function
    End If

    If rng.Count = 1 Then
        ' 单单元格: .Value 返回标量而非数组
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = rng.Value
    Else
        ' 多单元格: .Value 直接返回 1-based 2D Variant 数组
        result = rng.Value
    End If

    SelectionToArray2D = result
End Function

' 将 Variant 数组或标量写入以 startCell 为首单元格的区域。
' 自动根据数据大小调整输出范围。
' 支持: 2D 数组、1D 数组（默认输出为行）、标量。
' 参数 asColumn: True 时将 1D 数组输出为单列。
Public Sub ArrayToRange(ByRef data As Variant, _
                        ByRef startCell As Range, _
                        Optional ByVal asColumn As Boolean = False)
    Dim outArr() As Variant
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim is1D As Boolean
    Dim lb As Long, ub As Long
    Dim elemCount As Long
    Dim r0 As Long, c0 As Long

    If startCell Is Nothing Then Exit Sub
    If IsEmpty(data) Then Exit Sub

    If Not IsArray(data) Then
        startCell.Value = data
        Exit Sub
    End If

    ' 检测是否为 1D 数组
    Err.Clear
    On Error Resume Next
    nCols = UBound(data, 2)
    If Err.Number <> 0 Then
        is1D = True
        Err.Clear
    End If
    On Error GoTo 0

    If is1D Then
        lb = LBound(data)
        ub = UBound(data)
        elemCount = ub - lb + 1

        If asColumn Then
            nRows = elemCount
            nCols = 1
            ReDim outArr(1 To nRows, 1 To nCols)
            For i = 1 To nRows
                outArr(i, 1) = data(lb + i - 1)
            Next i
        Else
            nRows = 1
            nCols = elemCount
            ReDim outArr(1 To nRows, 1 To nCols)
            For i = 1 To nCols
                outArr(1, i) = data(lb + i - 1)
            Next i
        End If
    Else
        ' 2D 数组
        r0 = LBound(data, 1)
        c0 = LBound(data, 2)
        nRows = UBound(data, 1) - r0 + 1
        nCols = UBound(data, 2) - c0 + 1

        If r0 = 1 And c0 = 1 Then
            ' 已是 1-based，可直接写入
            startCell.Resize(nRows, nCols).Value = data
            Exit Sub
        End If

        ' 非 1-based，复制到 1-based 数组再写入
        ReDim outArr(1 To nRows, 1 To nCols)
        For i = 1 To nRows
            For j = 1 To nCols
                outArr(i, j) = data(r0 + i - 1, c0 + j - 1)
            Next j
        Next i
    End If

    startCell.Resize(nRows, nCols).Value = outArr
End Sub

'=============================================================================
' 基本矩阵操作
'=============================================================================
Public Function MatrixTranspose(ByRef A As Variant) As Double()
    ValidateMatrix A, "A"
    Dim m As Long, n As Long, i As Long, j As Long
    m = MatrixRows(A): n = MatrixCols(A)
    Dim result() As Double
    ReDim result(1 To n, 1 To m)
    For i = 1 To m
        For j = 1 To n
            result(j, i) = A(LBound(A, 1) + i - 1, LBound(A, 2) + j - 1)
        Next j
    Next i
    MatrixTranspose = result
End Function

Public Function IdentityMatrix(ByVal n As Long) As Double()
    If n < 1 Then
        Err.Raise ERR_IDENTITY_SIZE, "IdentityMatrix", "矩阵维度 n 必须 >= 1，实际为 " & n
    End If
    Dim result() As Double
    Dim i As Long
    ReDim result(1 To n, 1 To n)
    For i = 1 To n: result(i, i) = 1#: Next i
    IdentityMatrix = result
End Function

Public Function MatrixCopy(ByRef A As Variant) As Double()
    ValidateMatrix A, "A"
    Dim m As Long, n As Long, i As Long, j As Long
    m = MatrixRows(A): n = MatrixCols(A)
    Dim result() As Double
    ReDim result(1 To m, 1 To n)
    For i = 1 To m
        For j = 1 To n
            result(i, j) = A(LBound(A, 1) + i - 1, LBound(A, 2) + j - 1)
        Next j
    Next i
    MatrixCopy = result
End Function

Public Function MatrixGetColumn(ByRef A As Variant, ByVal k As Long) As Double()
    ValidateMatrix A, "A"
    Dim m As Long, n As Long, i As Long
    m = MatrixRows(A): n = MatrixCols(A)
    If k < 1 Or k > n Then
        Err.Raise ERR_COL_OUT_OF_RANGE, "MatrixGetColumn", _
            "列索引 k=" & k & " 越界，矩阵为 " & m & "x" & n
    End If
    Dim result() As Double
    ReDim result(1 To m)
    For i = 1 To m
        result(i) = A(LBound(A, 1) + i - 1, LBound(A, 2) + k - 1)
    Next i
    MatrixGetColumn = result
End Function

Public Sub MatrixSetColumn(ByRef A() As Double, ByVal k As Long, ByRef col() As Double)
    Dim m As Long, n As Long, i As Long
    m = MatrixRows(A): n = MatrixCols(A)
    If k < 1 Or k > n Then
        Err.Raise ERR_COL_LEN_MISMATCH, "MatrixSetColumn", _
            "列索引 k=" & k & " 越界，矩阵为 " & m & "x" & n
    End If
    If UBound(col) - LBound(col) + 1 <> m Then
        Err.Raise ERR_SETCOL_LEN, "MatrixSetColumn", _
            "列向量长度与矩阵行数不匹配: 向量 " & (UBound(col) - LBound(col) + 1) & ", 矩阵 " & m
    End If
    For i = 1 To m
        A(LBound(A, 1) + i - 1, LBound(A, 2) + k - 1) = col(LBound(col) + i - 1)
    Next i
End Sub

'=============================================================================
' 矩阵乘法
'=============================================================================
Public Function MatrixMultiply(ByRef A As Variant, _
                               ByRef B As Variant, _
                               Optional ByVal blockSize As Long = 32) As Double()
    Dim mA As Long, nA As Long, mB As Long, nB As Long
    Dim result() As Double
    Dim i As Long, j As Long, k As Long
    Dim ii As Long, kk As Long, jj As Long
    Dim i2 As Long, k2 As Long, j2 As Long
    Dim bs As Long
    Dim tempA As Double
    Dim rBk As Long
    Dim tmpSize As Long
    Dim rA0 As Long, cA0 As Long, rB0 As Long, cB0 As Long
    Dim tempRes() As Double

    ValidateMatrix A, "A"
    ValidateMatrix B, "B"
    If blockSize < 1 Then blockSize = 1

    mA = MatrixRows(A): nA = MatrixCols(A)
    mB = MatrixRows(B): nB = MatrixCols(B)
    If nA <> mB Then
        Err.Raise ERR_DIM_MISMATCH, "MatrixMultiply", _
            "维度不匹配: A为 " & mA & "x" & nA & ", B为 " & mB & "x" & nB
    End If
    ' 小矩阵使用朴素算法 (分块开销超过缓存收益; 阈值 64 按 SKILL.md §5.3)
    If nA <= 64 Or nB <= 64 Then
        MatrixMultiply = MatrixMultiplyNaive(A, B)
        Exit Function
    End If

    ReDim result(1 To mA, 1 To nB)
    rA0 = LBound(A, 1): cA0 = LBound(A, 2)
    rB0 = LBound(B, 1): cB0 = LBound(B, 2)
    bs = blockSize
    If bs > nA Then bs = nA

    ReDim tempRes(1 To bs)
    For ii = 1 To mA Step bs
        i2 = MinLong(ii + bs - 1, mA)
        For kk = 1 To nA Step bs
            k2 = MinLong(kk + bs - 1, nA)
            For jj = 1 To nB Step bs
                j2 = MinLong(jj + bs - 1, nB)
                tmpSize = j2 - jj + 1
                For i = ii To i2
                    For j = 1 To tmpSize
                        tempRes(j) = result(i, jj + j - 1)
                    Next j
                    For k = kk To k2
                        tempA = A(rA0 + i - 1, cA0 + k - 1)
                        If tempA <> 0# Then
                            rBk = rB0 + k - 1
                            For j = 1 To tmpSize
                                tempRes(j) = tempRes(j) + tempA * B(rBk, cB0 + jj + j - 2)
                            Next j
                        End If
                    Next k
                    For j = 1 To tmpSize
                        result(i, jj + j - 1) = tempRes(j)
                    Next j
                Next i
            Next jj
        Next kk
    Next ii

    MatrixMultiply = result
End Function

Public Function MatrixMultiplyNaive(ByRef A As Variant, ByRef B As Variant) As Double()
    Dim mA As Long, nA As Long, mB As Long, nB As Long
    Dim result() As Double
    Dim i As Long, j As Long, k As Long
    Dim rA0 As Long, cA0 As Long, rB0 As Long, cB0 As Long
    Dim s As Double, c As Double, y As Double, t As Double  ' Kahan compensation vars
    ValidateMatrix A, "A"
    ValidateMatrix B, "B"
    mA = MatrixRows(A): nA = MatrixCols(A)
    mB = MatrixRows(B): nB = MatrixCols(B)
    If nA <> mB Then Err.Raise ERR_DIM_MISMATCH, , "维度不匹配"

    ReDim result(1 To mA, 1 To nB)
    rA0 = LBound(A, 1)
    cA0 = LBound(A, 2)
    rB0 = LBound(B, 1)
    cB0 = LBound(B, 2)
    For i = 1 To mA
        For j = 1 To nB
            s = 0#: c = 0#  ' Kahan: sum and compensation
            For k = 1 To nA
                y = A(rA0 + i - 1, cA0 + k - 1) * B(rB0 + k - 1, cB0 + j - 1) - c
                t = s + y
                c = (t - s) - y  ' Kahan: recover lost low-order bits
                s = t
            Next k
            result(i, j) = s
        Next j
    Next i
    MatrixMultiplyNaive = result
End Function

'=============================================================================
' SVD 分解 (One-Sided Jacobi)
'=============================================================================
Public Sub SVD(ByRef A() As Double, _
               ByRef U() As Double, _
               ByRef S() As Double, _
               ByRef Vt() As Double, _
               Optional ByVal tol As Double = DEFAULT_TOL, _
               Optional ByVal maxSweeps As Long = MAX_SWEEPS)
    Dim m As Long, n As Long
    Dim At() As Double, U1() As Double, S1() As Double, Vt1() As Double
    ValidateMatrix A, "A"
    Dim workA() As Double
    workA = MatrixCopy(A)  ' 确保 1-based，内部函数依赖此约定，使用局部变量避免修改调用方数组
    m = MatrixRows(workA): n = MatrixCols(workA)

    If tol <= 0# Or tol < NUM_EPSILON Then tol = DEFAULT_TOL
    If maxSweeps < 1 Then maxSweeps = MAX_SWEEPS
    If maxSweeps > 100000 Then Err.Raise ERR_SWEEP_LIMIT, , "maxSweeps 过大"

    ' m < n (wide matrix): compute SVD of Aᵀ instead, then swap U ↔ Vt.
    ' A = U·Σ·Vt  ⇒  Aᵀ = V·Σ·Uᵀ  ⇒  SVD(Aᵀ) = (V, Σ, Uᵀ)
    ' So U = (Vt₁)ᵀ, S = S₁, Vt = (U₁)ᵀ
    If m < n Then
        At = MatrixTranspose(workA)
        JacobiWithScaling At, U1, S1, Vt1, tol, maxSweeps
        U = MatrixTranspose(Vt1)
        S = S1
        Vt = MatrixTranspose(U1)
        SortSVD U, S, Vt
        Exit Sub
    End If

    JacobiWithScaling workA, U, S, Vt, tol, maxSweeps
    SortSVD U, S, Vt
End Sub

' 对输入矩阵做缩放 + JacobiSVD，提取为独立过程以避免 SVD 中 m<n 分支重复代码
Private Sub JacobiWithScaling(ByRef A() As Double, _
                              ByRef U() As Double, _
                              ByRef S() As Double, _
                              ByRef Vt() As Double, _
                              ByVal tol As Double, _
                              ByVal maxSweeps As Long)
    Dim m As Long, n As Long, i As Long, j As Long
    m = MatrixRows(A): n = MatrixCols(A)

    Dim maxAbs As Double: maxAbs = 0#
    For i = 1 To m
        For j = 1 To n
            If Abs(A(i, j)) > maxAbs Then maxAbs = Abs(A(i, j))
        Next j
    Next i

    If maxAbs > 0# Then
        Dim Work() As Double
        ReDim Work(1 To m, 1 To n)
        For i = 1 To m
            For j = 1 To n
                Work(i, j) = A(i, j) / maxAbs
            Next j
        Next i
        JacobiSVD Work, U, S, Vt, tol, maxSweeps
        For i = 1 To MatrixRows(S): S(i, i) = S(i, i) * maxAbs: Next i
    Else
        JacobiSVD A, U, S, Vt, tol, maxSweeps
    End If
End Sub

Private Sub JacobiSVD(ByRef A() As Double, _
                      ByRef U() As Double, _
                      ByRef S() As Double, _
                      ByRef Vt() As Double, _
                      ByVal tol As Double, ByVal maxSweeps As Long)
    Dim m As Long, n As Long
    Dim sweep As Long, converged As Boolean
    Dim i As Long, j As Long, k As Long
    Dim c As Double, sn As Double, t As Double, zeta As Double
    Dim a_ii As Double, a_jj As Double, a_ij As Double
    Dim new_i As Double, new_j As Double
    Dim row As Long
    Dim urow_i As Double, urow_j As Double

    m = MatrixRows(A): n = MatrixCols(A)
    If tol <= 0# Or tol < NUM_EPSILON Then tol = DEFAULT_TOL
    If maxSweeps < 1 Then maxSweeps = MAX_SWEEPS

    Dim relTol As Double: relTol = tol
    Dim frob As Double
    frob = MatrixFrobeniusNorm(A)
    If frob > 0# Then relTol = tol * frob

    ' 标量情况
    If m = 1 And n = 1 Then
        ReDim U(1 To 1, 1 To 1)
        ReDim S(1 To 1, 1 To 1)
        ReDim Vt(1 To 1, 1 To 1)
        S(1, 1) = Abs(A(1, 1))
        U(1, 1) = 1#
        If A(1, 1) >= 0 Then Vt(1, 1) = 1# Else Vt(1, 1) = -1#
        Exit Sub
    End If

    U = MatrixCopy(A)
    ReDim Vt(1 To n, 1 To n)
    For i = 1 To n: Vt(i, i) = 1#: Next i
    ReDim S(1 To n, 1 To n)

    Dim tempColI() As Double, tempColJ() As Double
    ReDim tempColI(1 To m)
    ReDim tempColJ(1 To m)

    For sweep = 1 To maxSweeps
        converged = True
        For i = 1 To n - 1
            For j = i + 1 To n
                a_ii = 0#: a_jj = 0#: a_ij = 0#
                For row = 1 To m
                    tempColI(row) = U(row, i)
                    tempColJ(row) = U(row, j)
                    a_ii = a_ii + tempColI(row) ^ 2
                    a_jj = a_jj + tempColJ(row) ^ 2
                    a_ij = a_ij + tempColI(row) * tempColJ(row)
                Next row
                If Abs(a_ij) > relTol * Sqr(a_ii * a_jj) Then
                    converged = False
                    Dim denom As Double: denom = 2# * a_ij
                    ' 用乘法避免除法溢出：|a_jj-a_ii|/|denom| > 1E300 ⇔ |a_jj-a_ii| > 1E300*|denom|
                    If denom = 0# Or Abs(a_jj - a_ii) > 1E+300 * Abs(denom) Then
                        ' a_ij 极小 → 分母接近零 → zeta 溢出
                        ' 此时列为准正交，跳过旋转 (t ≈ 0)
                    Else
                        zeta = (a_jj - a_ii) / denom
                        ' 数值稳定: |zeta| 极大时 zeta*zeta 会溢出 Double
                        If Abs(zeta) < 1E+150 Then
                            t = SignNonZero(zeta) / (Abs(zeta) + Sqr(1# + zeta * zeta))
                        Else
                            t = 0.5 / zeta
                        End If
                        c = 1# / Sqr(1# + t * t): sn = c * t
                        For row = 1 To m
                            urow_i = tempColI(row)
                            urow_j = tempColJ(row)
                            U(row, i) = c * urow_i - sn * urow_j
                            U(row, j) = sn * urow_i + c * urow_j
                        Next row
                        For k = 1 To n
                            Dim vti_k As Double, vtj_k As Double
                            vti_k = Vt(i, k): vtj_k = Vt(j, k)
                            Vt(i, k) = c * vti_k - sn * vtj_k
                            Vt(j, k) = sn * vti_k + c * vtj_k
                        Next k
                    End If
                End If
            Next j
        Next i
        If converged Then Exit For
    Next sweep

    If Not converged Then
        Err.Raise ERR_SVD_CONVERGE, "SVD", _
            "SVD 未能在 " & maxSweeps & " 次扫描内收敛。矩阵可能病态。"
    End If

    For j = 1 To n
        a_jj = 0#
        For row = 1 To m
            a_jj = a_jj + U(row, j) ^ 2
        Next row
        a_jj = Sqr(a_jj)
        S(j, j) = a_jj
        If a_jj > NUM_EPSILON Then
            For row = 1 To m: U(row, j) = U(row, j) / a_jj: Next row
        Else
            For row = 1 To m: U(row, j) = 0#: Next row
        End If
    Next j
End Sub

Private Sub SortSVD(ByRef U() As Double, ByRef S() As Double, ByRef Vt() As Double)
    Dim k As Long, n As Long, i As Long, j As Long, bestJ As Long
    Dim tmp As Double, row As Long
    Dim uRows As Long
    k = MatrixRows(S)
    n = MatrixCols(Vt)
    For i = 1 To k - 1
        bestJ = i
        For j = i + 1 To k
            If S(j, j) > S(bestJ, bestJ) Then bestJ = j
        Next j
        If bestJ <> i Then
            tmp = S(i, i): S(i, i) = S(bestJ, bestJ): S(bestJ, bestJ) = tmp
            uRows = MatrixRows(U)
            For row = 1 To uRows
                tmp = U(row, i): U(row, i) = U(row, bestJ): U(row, bestJ) = tmp
            Next row
            For j = 1 To n
                tmp = Vt(i, j): Vt(i, j) = Vt(bestJ, j): Vt(bestJ, j) = tmp
            Next j
        End If
    Next i
End Sub

'=============================================================================
' 伪逆 & 矩阵秩
'=============================================================================

' 计算基于奇异值的容差，若最大奇异值为零则返回 False
Private Function SingularTolerance(ByRef S() As Double, _
                                   ByVal userTol As Double, _
                                   ByVal dimMax As Long, _
                                   ByRef maxSigma As Double, _
                                   ByRef tol As Double) As Boolean
    Dim k As Long, i As Long
    k = MatrixRows(S)
    maxSigma = 0#
    For i = 1 To k
        If S(i, i) > maxSigma Then maxSigma = S(i, i)
    Next i
    If maxSigma = 0# Then
        SingularTolerance = False
        Exit Function
    End If
    If userTol >= 0# Then
        tol = userTol
    Else
        tol = maxSigma * NUM_EPSILON * CDbl(dimMax)
    End If
    SingularTolerance = True
End Function

Public Function PseudoInverse(ByRef A As Variant, _
                              Optional ByVal tolerance As Double = -1#) As Double()
    Dim matA() As Double: matA = ToDoubleMatrix(A)
    Dim k As Long, i As Long
    Dim U() As Double, S() As Double, Vt() As Double
    Dim V() As Double, Ut() As Double, sigmaP() As Double
    Dim temp() As Double, result() As Double

    SVD matA, U, S, Vt  ' ValidateMatrix 在 SVD 内部调用
    k = MatrixRows(S)

    Dim maxSigma As Double, tol As Double
    Dim dimMax As Long
    If MatrixRows(matA) > MatrixCols(matA) Then dimMax = MatrixRows(matA) Else dimMax = MatrixCols(matA)
    If Not SingularTolerance(S, tolerance, dimMax, maxSigma, tol) Then
        ReDim result(1 To MatrixCols(A), 1 To MatrixRows(A))
        PseudoInverse = result
        Exit Function
    End If

    ReDim sigmaP(1 To k, 1 To k)
    For i = 1 To k
        If S(i, i) > tol Then sigmaP(i, i) = 1# / S(i, i)
    Next i

    V = MatrixTranspose(Vt)
    Ut = MatrixTranspose(U)
    temp = MatrixMultiply(V, sigmaP)
    result = MatrixMultiply(temp, Ut)
    PseudoInverse = result
End Function

Public Function MatrixRank_Array(ByRef A As Variant, _
                                 Optional ByVal tolerance As Double = -1#) As Long
    Dim U() As Double, S() As Double, Vt() As Double
    Dim matA() As Double: matA = ToDoubleMatrix(A)
    Dim k As Long, i As Long, rnk As Long
    SVD matA, U, S, Vt  ' ValidateMatrix 在 SVD 内部调用
    k = MatrixRows(S)

    Dim maxSigma As Double, tol As Double
    Dim dimMax As Long
    If MatrixRows(matA) > MatrixCols(matA) Then dimMax = MatrixRows(matA) Else dimMax = MatrixCols(matA)
    If Not SingularTolerance(S, tolerance, dimMax, maxSigma, tol) Then
        MatrixRank_Array = 0
        Exit Function
    End If

    For i = 1 To k
        If S(i, i) > tol Then rnk = rnk + 1
    Next i
    MatrixRank_Array = rnk
End Function

'=============================================================================
' 对称矩阵特征分解 (Jacobi)
'=============================================================================
Public Sub EigenSymmetric(ByRef A() As Double, _
                          ByRef V() As Double, _
                          ByRef D() As Double, _
                          Optional ByVal tol As Double = DEFAULT_TOL, _
                          Optional ByVal maxSweeps As Long = MAX_SWEEPS)
    Dim n As Long, sweep As Long, converged As Boolean
    Dim p As Long, qIdx As Long
    Dim theta As Double, t As Double, c As Double, s As Double
    Dim a_pp As Double, a_qq As Double, a_pq As Double
    Dim i As Long, j As Long, bestJ As Long
    Dim new_ip As Double, new_iq As Double
    Dim maxAbsA As Double
    Dim symTol As Double
    Dim tmp As Double

    ValidateMatrix A, "A"
    Dim workA() As Double
    workA = MatrixCopy(A)  ' 确保 1-based，内部函数依赖此约定，使用局部变量避免修改调用方数组
    n = MatrixRows(workA)
    If MatrixCols(workA) <> n Then Err.Raise ERR_NOT_SQUARE, , "矩阵必须是方阵。"
    If tol <= 0# Or tol < NUM_EPSILON Then tol = DEFAULT_TOL
    If maxSweeps < 1 Then maxSweeps = MAX_SWEEPS
    If maxSweeps > 100000 Then Err.Raise ERR_SWEEP_LIMIT, , "maxSweeps 过大"

    maxAbsA = 0#
    For i = 1 To n
        For j = 1 To n
            If Abs(workA(i, j)) > maxAbsA Then maxAbsA = Abs(workA(i, j))
        Next j
    Next i
    symTol = maxAbsA * NUM_EPSILON * CDbl(n)
    If symTol < DEFAULT_TOL Then symTol = DEFAULT_TOL

    ' 合并对称性检查和输入缩放到一次遍历
    If maxAbsA > 0# Then
        ReDim D(1 To n, 1 To n)
        For i = 1 To n
            For j = 1 To n
                D(i, j) = workA(i, j) / maxAbsA
                If j > i Then
                    If Abs(workA(i, j) - workA(j, i)) > symTol Then
                        Err.Raise ERR_NOT_SYMMETRIC, , "矩阵不对称，请使用 SVD。"
                    End If
                End If
            Next j
        Next i
    Else
        D = MatrixCopy(workA)
    End If

    V = IdentityMatrix(n)

    For sweep = 1 To maxSweeps
        converged = True
        For p = 1 To n - 1
            For qIdx = p + 1 To n
                a_pq = D(p, qIdx)
                If Abs(a_pq) > tol * (Abs(D(p, p)) + Abs(D(qIdx, qIdx)) + NUM_EPSILON) Then
                    converged = False
                    a_pp = D(p, p): a_qq = D(qIdx, qIdx)
                    theta = (a_qq - a_pp) / (2# * a_pq)
                    If Abs(theta) < 1E+15 Then
                        t = SignNonZero(theta) / (Abs(theta) + Sqr(1# + theta * theta))
                    Else
                        t = 0.5 / theta
                    End If
                    c = 1# / Sqr(1# + t * t): s = c * t

                    For i = 1 To n
                        If i <> p And i <> qIdx Then
                            new_ip = c * D(i, p) - s * D(i, qIdx)
                            new_iq = s * D(i, p) + c * D(i, qIdx)
                            D(i, p) = new_ip: D(i, qIdx) = new_iq
                            D(p, i) = new_ip: D(qIdx, i) = new_iq
                        End If
                    Next i
                    D(p, p) = c * c * a_pp + s * s * a_qq - 2# * c * s * a_pq
                    D(qIdx, qIdx) = s * s * a_pp + c * c * a_qq + 2# * c * s * a_pq
                    D(p, qIdx) = 0#: D(qIdx, p) = 0#

                    For i = 1 To n
                        new_ip = c * V(i, p) - s * V(i, qIdx)
                        new_iq = s * V(i, p) + c * V(i, qIdx)
                        V(i, p) = new_ip: V(i, qIdx) = new_iq
                    Next i
                End If
            Next qIdx
        Next p
        If converged Then Exit For
    Next sweep

    If Not converged Then
        Err.Raise ERR_EIGEN_CONVERGE, "EigenSymmetric", _
            "特征分解未能在 " & maxSweeps & " 次扫描内收敛。矩阵可能病态。"
    End If

    ' 恢复缩放
    If maxAbsA > 0# Then
        For i = 1 To n: D(i, i) = D(i, i) * maxAbsA: Next i
    End If

    ' 特征值降序排序
    For i = 1 To n - 1
        bestJ = i
        For j = i + 1 To n
            If D(j, j) > D(bestJ, bestJ) Then bestJ = j
        Next j
        If bestJ <> i Then
            tmp = D(i, i): D(i, i) = D(bestJ, bestJ): D(bestJ, bestJ) = tmp
            For j = 1 To n
                tmp = V(j, i): V(j, i) = V(j, bestJ): V(j, bestJ) = tmp
            Next j
        End If
    Next i
End Sub

'=============================================================================
' QR 分解 (Householder)
'=============================================================================
Public Sub QRDecomposition(ByRef A() As Double, _
                           ByRef Q() As Double, _
                           ByRef R() As Double, _
                           Optional ByVal economy As Boolean = False)
    Dim m As Long, n As Long, k As Long, maxK As Long
    Dim i As Long, j As Long, row As Long
    Dim u() As Double, beta As Double, v As Double
    Dim ulen As Long, xnorm As Double, colMax As Double

    ValidateMatrix A, "A"
    m = MatrixRows(A): n = MatrixCols(A)
    If m > 10000 Or n > 10000 Then Err.Raise ERR_MATRIX_TOO_BIG, , "矩阵过大。"
    If m < 1 Or n < 1 Then Err.Raise ERR_DIM_INVALID, , "矩阵维度无效。"

    R = MatrixCopy(A)
    Q = IdentityMatrix(m)

    maxK = MinLong(m, n)

    Dim rR0 As Long: rR0 = LBound(R, 1)
    Dim cR0 As Long: cR0 = LBound(R, 2)
    Dim rQ0 As Long: rQ0 = LBound(Q, 1)
    Dim cQ0 As Long: cQ0 = LBound(Q, 2)

    ReDim u(1 To m)
    For k = 1 To maxK
        ulen = m - k + 1
        xnorm = 0#: colMax = 0#
        For i = k To m
            u(i - k + 1) = R(rR0 + i - 1, cR0 + k - 1)
            xnorm = xnorm + u(i - k + 1) ^ 2
            If Abs(u(i - k + 1)) > colMax Then colMax = Abs(u(i - k + 1))
        Next i
        xnorm = Sqr(xnorm)
        ' Relative tolerance per SKILL.md S8: Abs(pivot) < 1E-12 * columnNorm
        If xnorm > NUM_EPSILON * colMax Then
            If u(1) > 0 Then u(1) = u(1) + xnorm Else u(1) = u(1) - xnorm
            beta = 0#
            For i = 1 To ulen: beta = beta + u(i) ^ 2: Next i
            beta = 2# / beta

            ' 更新 R (左乘 H)
            For j = k To n
                v = 0#
                For i = 1 To ulen
                    v = v + u(i) * R(rR0 + k + i - 2, cR0 + j - 1)
                Next i
                v = v * beta
                For i = 1 To ulen
                    R(rR0 + k + i - 2, cR0 + j - 1) = _
                        R(rR0 + k + i - 2, cR0 + j - 1) - v * u(i)
                Next i
            Next j

            ' 更新 Q (右乘 H^T)
            For row = 1 To m
                v = 0#
                For i = 1 To ulen
                    v = v + u(i) * Q(rQ0 + row - 1, cQ0 + k + i - 2)
                Next i
                v = v * beta
                For i = 1 To ulen
                    Q(rQ0 + row - 1, cQ0 + k + i - 2) = _
                        Q(rQ0 + row - 1, cQ0 + k + i - 2) - v * u(i)
                Next i
            Next row
        End If
    Next k

    ' 经济模式
    If economy Then
        Dim Rout() As Double, Qout() As Double
        ' 确保 R 方阵: maxK x maxK (只保留三角部分)
        ReDim Rout(1 To maxK, 1 To maxK)
        Dim rr As Long, cc As Long
        For rr = 1 To maxK
            For cc = 1 To maxK
                If rr <= cc Then
                    Rout(rr, cc) = R(rR0 + rr - 1, cR0 + cc - 1)
                Else
                    Rout(rr, cc) = 0#
                End If
            Next cc
        Next rr
        ReDim Qout(1 To m, 1 To maxK)
        For rr = 1 To m
            For cc = 1 To maxK
                Qout(rr, cc) = Q(rQ0 + rr - 1, cQ0 + cc - 1)
            Next cc
        Next rr
        R = Rout: Q = Qout
    End If
End Sub

'=============================================================================
' QRDecompositionPiv — 带列主元的 QR 分解 (Businger-Golub)
'
' 计算 A*P = Q*R, 其中 P 为列置换矩阵
' perm(j) 返回原始列索引: 第 k 个主元来自原始第 perm(k) 列
' 用于秩亏检测、最小二乘、子集选择
'=============================================================================
Public Sub QRDecompositionPiv(ByRef A() As Double, _
                              ByRef Q() As Double, _
                              ByRef R() As Double, _
                              ByRef perm() As Long, _
                              Optional ByVal economy As Boolean = False)
    Dim m As Long, n As Long, k As Long, maxK As Long
    Dim i As Long, j As Long, row As Long
    Dim u() As Double, beta As Double, v As Double
    Dim ulen As Long, xnorm As Double, colMax As Double

    ValidateMatrix A, "A"
    m = MatrixRows(A): n = MatrixCols(A)
    If m > 10000 Or n > 10000 Then Err.Raise ERR_MATRIX_TOO_BIG, , "矩阵过大。"
    If m < 1 Or n < 1 Then Err.Raise ERR_DIM_INVALID, , "矩阵维度无效。"

    R = MatrixCopy(A)
    Q = IdentityMatrix(m)
    maxK = MinLong(m, n)

    Dim rR0 As Long: rR0 = LBound(R, 1)
    Dim cR0 As Long: cR0 = LBound(R, 2)
    Dim rQ0 As Long: rQ0 = LBound(Q, 1)
    Dim cQ0 As Long: cQ0 = LBound(Q, 2)

    ' --- Column norm bookkeeping ---
    Dim colNorms() As Double
    ReDim perm(1 To n)
    ReDim colNorms(1 To n)
    For j = 1 To n
        perm(j) = j
        Dim cn As Double: cn = 0#
        For i = 1 To m
            cn = cn + R(rR0 + i - 1, cR0 + j - 1) ^ 2
        Next i
        colNorms(j) = cn
    Next j

    ReDim u(1 To m)
    For k = 1 To maxK
        ' --- Pivot: pick column with largest residual norm ---
        Dim maxNorm As Double: maxNorm = -1#
        Dim pivotCol As Long: pivotCol = k
        For j = k To n
            If colNorms(j) > maxNorm Then
                maxNorm = colNorms(j)
                pivotCol = j
            End If
        Next j
        If pivotCol <> k Then
            Dim tmpSwap As Double
            For i = 1 To m
                tmpSwap = R(rR0 + i - 1, cR0 + k - 1)
                R(rR0 + i - 1, cR0 + k - 1) = R(rR0 + i - 1, cR0 + pivotCol - 1)
                R(rR0 + i - 1, cR0 + pivotCol - 1) = tmpSwap
            Next i
            Dim tmpD As Double: tmpD = colNorms(k): colNorms(k) = colNorms(pivotCol): colNorms(pivotCol) = tmpD
            Dim tmpL As Long: tmpL = perm(k): perm(k) = perm(pivotCol): perm(pivotCol) = tmpL
        End If

        ulen = m - k + 1
        xnorm = 0#: colMax = 0#
        For i = k To m
            u(i - k + 1) = R(rR0 + i - 1, cR0 + k - 1)
            xnorm = xnorm + u(i - k + 1) ^ 2
            If Abs(u(i - k + 1)) > colMax Then colMax = Abs(u(i - k + 1))
        Next i
        xnorm = Sqr(xnorm)
        If xnorm > NUM_EPSILON * colMax Then
            If u(1) > 0 Then u(1) = u(1) + xnorm Else u(1) = u(1) - xnorm
            beta = 0#
            For i = 1 To ulen: beta = beta + u(i) ^ 2: Next i
            beta = 2# / beta

            For j = k To n
                v = 0#
                For i = 1 To ulen
                    v = v + u(i) * R(rR0 + k + i - 2, cR0 + j - 1)
                Next i
                v = v * beta
                For i = 1 To ulen
                    R(rR0 + k + i - 2, cR0 + j - 1) = _
                        R(rR0 + k + i - 2, cR0 + j - 1) - v * u(i)
                Next i
            Next j

            For row = 1 To m
                v = 0#
                For i = 1 To ulen
                    v = v + u(i) * Q(rQ0 + row - 1, cQ0 + k + i - 2)
                Next i
                v = v * beta
                For i = 1 To ulen
                    Q(rQ0 + row - 1, cQ0 + k + i - 2) = _
                        Q(rQ0 + row - 1, cQ0 + k + i - 2) - v * u(i)
                Next i
            Next row

            ' Downdate residual column norms
            For j = k + 1 To n
                Dim rKJ As Double: rKJ = R(rR0 + k - 1, cR0 + j - 1)
                colNorms(j) = colNorms(j) - rKJ * rKJ
                If colNorms(j) < 0# Then colNorms(j) = 0#
            Next j
        End If
    Next k

    ' 经济模式
    If economy Then
        Dim Rout() As Double, Qout() As Double
        ReDim Rout(1 To maxK, 1 To maxK)
        Dim rr As Long, cc As Long
        For rr = 1 To maxK
            For cc = 1 To maxK
                If rr <= cc Then
                    Rout(rr, cc) = R(rR0 + rr - 1, cR0 + cc - 1)
                Else
                    Rout(rr, cc) = 0#
                End If
            Next cc
        Next rr
        ReDim Qout(1 To m, 1 To maxK)
        For rr = 1 To m
            For cc = 1 To maxK
                Qout(rr, cc) = Q(rQ0 + rr - 1, cQ0 + cc - 1)
            Next cc
        Next rr
        R = Rout: Q = Qout
    End If
End Sub

'=============================================================================
' 工作表函数 (UDF_LINALG_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_LINALG_SVD_U(ByVal rng As Variant, Optional ByVal tol As Variant = DEFAULT_TOL, _
                      Optional ByVal maxSweeps As Variant = MAX_SWEEPS) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, U() As Double, S() As Double, Vt() As Double
    A = RangeToMatrix(rng)
    SVD A, U, S, Vt, tol, maxSweeps
    UDF_LINALG_SVD_U = DblMatrixToVariant(U)
    Exit Function
ErrHandler:
    UDF_LINALG_SVD_U = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_SVD_S(ByVal rng As Variant, Optional ByVal tol As Variant = DEFAULT_TOL, _
                      Optional ByVal maxSweeps As Variant = MAX_SWEEPS) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, U() As Double, S() As Double, Vt() As Double
    A = RangeToMatrix(rng)
    SVD A, U, S, Vt, tol, maxSweeps
    UDF_LINALG_SVD_S = DblMatrixToVariant(S)
    Exit Function
ErrHandler:
    UDF_LINALG_SVD_S = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_SVD_VT(ByVal rng As Variant, Optional ByVal tol As Variant = DEFAULT_TOL, _
                       Optional ByVal maxSweeps As Variant = MAX_SWEEPS) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, U() As Double, S() As Double, Vt() As Double
    A = RangeToMatrix(rng)
    SVD A, U, S, Vt, tol, maxSweeps
    UDF_LINALG_SVD_VT = DblMatrixToVariant(Vt)
    Exit Function
ErrHandler:
    UDF_LINALG_SVD_VT = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_SVD_SVALS(ByVal rng As Variant, Optional ByVal tol As Variant = DEFAULT_TOL, _
                          Optional ByVal maxSweeps As Variant = MAX_SWEEPS) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, U() As Double, S() As Double, Vt() As Double
    Dim i As Long, k As Long
    A = RangeToMatrix(rng)
    SVD A, U, S, Vt, tol, maxSweeps
    k = MatrixRows(S)
    Dim result() As Variant
    ReDim result(1 To k, 1 To 1)
    For i = 1 To k: result(i, 1) = S(i, i): Next i
    UDF_LINALG_SVD_SVALS = result
    Exit Function
ErrHandler:
    UDF_LINALG_SVD_SVALS = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_PINV(ByVal rng As Variant, Optional ByVal tolerance As Variant = -1#) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, Ap() As Double
    A = RangeToMatrix(rng)
    Ap = PseudoInverse(A, tolerance)
    UDF_LINALG_PINV = DblMatrixToVariant(Ap)
    Exit Function
ErrHandler:
    UDF_LINALG_PINV = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_QR_Q(ByVal rng As Variant, Optional ByVal economy As Variant = False) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, Q() As Double, R() As Double
    A = RangeToMatrix(rng)
    QRDecomposition A, Q, R, economy
    UDF_LINALG_QR_Q = DblMatrixToVariant(Q)
    Exit Function
ErrHandler:
    UDF_LINALG_QR_Q = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_QR_R(ByVal rng As Variant, Optional ByVal economy As Variant = False) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, Q() As Double, R() As Double
    A = RangeToMatrix(rng)
    QRDecomposition A, Q, R, economy
    UDF_LINALG_QR_R = DblMatrixToVariant(R)
    Exit Function
ErrHandler:
    UDF_LINALG_QR_R = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_EIGVAL(ByVal rng As Variant, Optional ByVal tol As Variant = DEFAULT_TOL, _
                           Optional ByVal maxSweeps As Variant = MAX_SWEEPS) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, V() As Double, D() As Double
    A = RangeToMatrix(rng)
    EigenSymmetric A, V, D, tol, maxSweeps
    UDF_LINALG_EIGVAL = DblMatrixToVariant(D)
    Exit Function
ErrHandler:
    UDF_LINALG_EIGVAL = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_EIGVEC(ByVal rng As Variant, Optional ByVal tol As Variant = DEFAULT_TOL, _
                           Optional ByVal maxSweeps As Variant = MAX_SWEEPS) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, V() As Double, D() As Double
    A = RangeToMatrix(rng)
    EigenSymmetric A, V, D, tol, maxSweeps
    UDF_LINALG_EIGVEC = DblMatrixToVariant(V)
    Exit Function
ErrHandler:
    UDF_LINALG_EIGVEC = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_RANK(ByVal rng As Variant, Optional ByVal tolerance As Variant = -1#) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double
    A = RangeToMatrix(rng)
    UDF_LINALG_RANK = MatrixRank_Array(A, tolerance)
    Exit Function
ErrHandler:
    UDF_LINALG_RANK = CVErr(xlErrValue)
End Function

'=============================================================================
' 矩阵算术运算
'=============================================================================

' MatrixTrace — 方阵的迹 (对角线元素之和)
Public Function MatrixTrace(ByRef A As Variant) As Double
    ValidateMatrix A, "A"
    Dim m As Long, n As Long, i As Long
    m = MatrixRows(A): n = MatrixCols(A)
    If m <> n Then Err.Raise ERR_NOT_SQUARE, "MatrixTrace", "矩阵必须是方阵。"

    Dim r0 As Long: r0 = LBound(A, 1)
    Dim c0 As Long: c0 = LBound(A, 2)
    Dim result As Double: result = 0#
    For i = 1 To n
        result = result + A(r0 + i - 1, c0 + i - 1)
    Next i
    MatrixTrace = result
End Function

' MatrixScale — 标量乘法
Public Function MatrixScale(ByRef A As Variant, ByVal scalar As Double) As Double()
    ValidateMatrix A, "A"
    Dim m As Long, n As Long, i As Long, j As Long
    m = MatrixRows(A): n = MatrixCols(A)
    Dim result() As Double
    ReDim result(1 To m, 1 To n)
    Dim r0 As Long: r0 = LBound(A, 1)
    Dim c0 As Long: c0 = LBound(A, 2)
    For i = 1 To m
        For j = 1 To n
            result(i, j) = A(r0 + i - 1, c0 + j - 1) * scalar
        Next j
    Next i
    MatrixScale = result
End Function

' MatrixAdd — 矩阵加法
Public Function MatrixAdd(ByRef A As Variant, ByRef B As Variant) As Double()
    ValidateMatrix A, "A"
    ValidateMatrix B, "B"
    Dim mA As Long, nA As Long, mB As Long, nB As Long
    mA = MatrixRows(A): nA = MatrixCols(A)
    mB = MatrixRows(B): nB = MatrixCols(B)
    If mA <> mB Or nA <> nB Then
        Err.Raise ERR_ADD_MISMATCH, "MatrixAdd", "矩阵维度不匹配: A=" & mA & "x" & nA & ", B=" & mB & "x" & nB
    End If

    Dim result() As Double
    ReDim result(1 To mA, 1 To nA)
    Dim i As Long, j As Long
    Dim rA0 As Long: rA0 = LBound(A, 1): Dim cA0 As Long: cA0 = LBound(A, 2)
    Dim rB0 As Long: rB0 = LBound(B, 1): Dim cB0 As Long: cB0 = LBound(B, 2)
    For i = 1 To mA
        For j = 1 To nA
            result(i, j) = A(rA0 + i - 1, cA0 + j - 1) + B(rB0 + i - 1, cB0 + j - 1)
        Next j
    Next i
    MatrixAdd = result
End Function

' MatrixSubtract — 矩阵减法 (A - B)
Public Function MatrixSubtract(ByRef A As Variant, ByRef B As Variant) As Double()
    ValidateMatrix A, "A"
    ValidateMatrix B, "B"
    Dim mA As Long, nA As Long, mB As Long, nB As Long
    mA = MatrixRows(A): nA = MatrixCols(A)
    mB = MatrixRows(B): nB = MatrixCols(B)
    If mA <> mB Or nA <> nB Then
        Err.Raise ERR_SUB_MISMATCH, "MatrixSubtract", "矩阵维度不匹配: A=" & mA & "x" & nA & ", B=" & mB & "x" & nB
    End If

    Dim result() As Double
    ReDim result(1 To mA, 1 To nA)
    Dim i As Long, j As Long
    Dim rA0 As Long: rA0 = LBound(A, 1): Dim cA0 As Long: cA0 = LBound(A, 2)
    Dim rB0 As Long: rB0 = LBound(B, 1): Dim cB0 As Long: cB0 = LBound(B, 2)
    For i = 1 To mA
        For j = 1 To nA
            result(i, j) = A(rA0 + i - 1, cA0 + j - 1) - B(rB0 + i - 1, cB0 + j - 1)
        Next j
    Next i
    MatrixSubtract = result
End Function

' MatrixHadamard — 哈达玛 (逐元素) 乘积
Public Function MatrixHadamard(ByRef A As Variant, ByRef B As Variant) As Double()
    ValidateMatrix A, "A"
    ValidateMatrix B, "B"
    Dim mA As Long, nA As Long, mB As Long, nB As Long
    mA = MatrixRows(A): nA = MatrixCols(A)
    mB = MatrixRows(B): nB = MatrixCols(B)
    If mA <> mB Or nA <> nB Then
        Err.Raise ERR_HADAMARD_MISMATCH, "MatrixHadamard", "矩阵维度不匹配。"
    End If

    Dim result() As Double
    ReDim result(1 To mA, 1 To nA)
    Dim i As Long, j As Long
    Dim rA0 As Long: rA0 = LBound(A, 1): Dim cA0 As Long: cA0 = LBound(A, 2)
    Dim rB0 As Long: rB0 = LBound(B, 1): Dim cB0 As Long: cB0 = LBound(B, 2)
    For i = 1 To mA
        For j = 1 To nA
            result(i, j) = A(rA0 + i - 1, cA0 + j - 1) * B(rB0 + i - 1, cB0 + j - 1)
        Next j
    Next i
    MatrixHadamard = result
End Function

' MatrixNorm — 矩阵范数
' normType: "1" = 1-范数 (列和最大), "inf" = Z_SYM范数 (行和最大), "fro" = Frobenius 范数 (默认)
Public Function MatrixNorm(ByRef A As Variant, Optional ByVal normType As String = "fro") As Double
    ValidateMatrix A, "A"
    Dim m As Long, n As Long, i As Long, j As Long
    m = MatrixRows(A): n = MatrixCols(A)

    Select Case LCase$(normType)
        Case "1"
            Dim colSum As Double, maxColSum As Double
            maxColSum = 0#
            For j = 1 To n
                colSum = 0#
                For i = 1 To m
                    colSum = colSum + Abs(A(LBound(A, 1) + i - 1, LBound(A, 2) + j - 1))
                Next i
                If colSum > maxColSum Then maxColSum = colSum
            Next j
            MatrixNorm = maxColSum
        Case "inf", "infinity"
            Dim rowSum As Double, maxRowSum As Double
            maxRowSum = 0#
            For i = 1 To m
                rowSum = 0#
                For j = 1 To n
                    rowSum = rowSum + Abs(A(LBound(A, 1) + i - 1, LBound(A, 2) + j - 1))
                Next j
                If rowSum > maxRowSum Then maxRowSum = rowSum
            Next i
            MatrixNorm = maxRowSum
        Case Else ' "fro"
            MatrixNorm = MatrixFrobeniusNorm(A)
    End Select
End Function

' MatrixPower — 方阵的 n 次幂 (n >= 0，使用二分幂运算)
Public Function MatrixPower(ByRef A As Variant, ByVal n As Long) As Double()
    ValidateMatrix A, "A"
    Dim d As Long: d = MatrixRows(A)
    If MatrixCols(A) <> d Then Err.Raise ERR_POWER_SQUARE, "MatrixPower", "矩阵必须是方阵。"
    If n < 0 Then Err.Raise ERR_NEG_EXPONENT, "MatrixPower", "负指数不支持。请对逆矩阵使用正整数幂。"

    If n = 0 Then
        MatrixPower = IdentityMatrix(d)
        Exit Function
    End If

    ' 二分幂运算 (exponentiation by squaring).
    ' 注意: 大指数可能导致 Double 溢出 (≈1E308), 无硬性上限但需注意数值范围.
    Dim result() As Double: result = IdentityMatrix(d)
    Dim base() As Double: base = MatrixCopy(A)
    Dim pow As Long: pow = n
    Do While pow > 0
        If (pow And 1) Then result = MatrixMultiply(result, base)
        pow = pow \ 2
        If pow > 0 Then base = MatrixMultiply(base, base)
    Loop
    MatrixPower = result
End Function

'=============================================================================
' LU 分解 (Doolittle, 部分主元)
'
' A = P * L * U
' L: 单位下三角 (对角线为 1)
' U: 上三角
' P: 置换矩阵 (行交换)
'
' 对于奇异矩阵, 若某主元为零则报错
'=============================================================================
Public Sub LUDecomposition(ByRef A() As Double, _
                            ByRef L() As Double, _
                            ByRef U() As Double, _
                            ByRef P() As Double, _
                            Optional ByRef outSwapCount As Long)

    Dim n As Long, i As Long, j As Long, k As Long
    Dim maxVal As Double, absVal As Double, pivVal As Double
    Dim pivot As Long, tmp As Double
    Dim swapCount As Long

    ValidateMatrix A, "A"
    Dim work() As Double: work = MatrixCopy(A)
    n = MatrixRows(work)
    If MatrixCols(work) <> n Then Err.Raise ERR_LU_SQUARE, "LUDecomposition", "矩阵必须是方阵。"

    ' 初始化 P = I
    P = IdentityMatrix(n)
    ReDim L(1 To n, 1 To n)
    ReDim U(1 To n, 1 To n)

    ' 计算相对容差 (基于矩阵最大绝对值)
    Dim pivotTol As Double: pivotTol = 0#
    For i = 1 To n
        For j = 1 To n
            If Abs(work(i, j)) > pivotTol Then pivotTol = Abs(work(i, j))
        Next j
    Next i
    pivotTol = pivotTol * NUM_EPSILON * CDbl(n)
    If pivotTol < NUM_EPSILON Then pivotTol = NUM_EPSILON

    For k = 1 To n
        ' 部分主元: 在第 k 列中找绝对值最大的行
        ' 注意: 必须基于消元后的列值 (实质 = work(i,k) - Σ_{m=1}^{k-1} L(i,m)*U(m,k))
        maxVal = 0#: pivot = k
        For i = k To n
            pivVal = work(i, k)
            For j = 1 To k - 1
                pivVal = pivVal - L(i, j) * U(j, k)
            Next j
            absVal = Abs(pivVal)
            If absVal > maxVal Then
                maxVal = absVal
                pivot = i
            End If
        Next i
        If maxVal < pivotTol Then
            Err.Raise ERR_SINGULAR, "LUDecomposition", "矩阵奇异或接近奇异，无法进行 LU 分解。"
        End If

        ' 交换行
        If pivot <> k Then
            swapCount = swapCount + 1
            For j = 1 To n
                tmp = work(k, j): work(k, j) = work(pivot, j): work(pivot, j) = tmp
                tmp = P(k, j): P(k, j) = P(pivot, j): P(pivot, j) = tmp
            Next j
            ' 同时交换 L 中已计算的乘数 (列 1..k-1)
            If k > 1 Then
                For j = 1 To k - 1
                    tmp = L(k, j): L(k, j) = L(pivot, j): L(pivot, j) = tmp
                Next j
            End If
        End If

        ' 计算 U(k, j) 和 L(i, k)
        For j = k To n
            U(k, j) = work(k, j)
            For i = 1 To k - 1
                U(k, j) = U(k, j) - L(k, i) * U(i, j)
            Next i
        Next j

        L(k, k) = 1#
        For i = k + 1 To n
            L(i, k) = work(i, k)
            For j = 1 To k - 1
                L(i, k) = L(i, k) - L(i, j) * U(j, k)
            Next j
            If Abs(U(k, k)) < pivotTol Then
                Err.Raise ERR_SINGULAR, "LUDecomposition", _
                    "主元为零 (pivot[" & k & "]=" & U(k, k) & ")。矩阵奇异。"
            End If
            L(i, k) = L(i, k) / U(k, k)
        Next i
    Next k
    outSwapCount = swapCount
End Sub

'=============================================================================
' CholeskyDecomposition — Cholesky 分解 (A = L * Lᵀ)
'
' 要求: A 为对称正定矩阵
'   L: 下三角矩阵 (单位: 1-based)
'   若矩阵非正定则报错
'
' 算法: 外积 Cholesky (列优先), O(n³/3)
'=============================================================================
Public Sub CholeskyDecomposition(ByRef A() As Double, ByRef L() As Double)
    Dim n As Long, j As Long, k As Long, i As Long
    Dim s As Double, tmp As Double
    Dim maxDiag As Double

    ValidateMatrix A, "A"
    n = MatrixRows(A)
    If MatrixCols(A) <> n Then Err.Raise ERR_CHOLESKY_SIZE, "CholeskyDecomposition", "矩阵必须是方阵。"

    ' Compute relative tolerance based on matrix scale
    Dim maxAbsSym As Double: maxAbsSym = 0#
    Dim symI As Long, symJ As Long
    For symI = 1 To n
        For symJ = 1 To n
            Dim asi As Double: asi = Abs(A(LBound(A, 1) + symI - 1, LBound(A, 2) + symJ - 1))
            If asi > maxAbsSym Then maxAbsSym = asi
        Next symJ
    Next symI
    ' 对称性检查: A(i,j) 必须等于 A(j,i)
    ' Use relative tolerance to handle large-magnitude matrices
    If maxAbsSym > 0# Then
        Dim symTol As Double: symTol = maxAbsSym * NUM_EPSILON * CDbl(n)
        If symTol < DEFAULT_TOL Then symTol = DEFAULT_TOL
        For symI = 1 To n
            For symJ = symI + 1 To n
                If Abs(A(LBound(A, 1) + symI - 1, LBound(A, 2) + symJ - 1) - _
                       A(LBound(A, 1) + symJ - 1, LBound(A, 2) + symI - 1)) > symTol Then
                    Err.Raise ERR_CHOLESKY_NOT_SYM, "CholeskyDecomposition", "矩阵非对称。"
                End If
            Next symJ
        Next symI
    End If

    ReDim L(1 To n, 1 To n)

    ' 计算 A 对角元最大绝对值作为正定性容差参考
    maxDiag = 0#
    Dim diagIdx As Long
    For diagIdx = 1 To n
        If Abs(A(LBound(A, 1) + diagIdx - 1, LBound(A, 2) + diagIdx - 1)) > maxDiag Then
            maxDiag = Abs(A(LBound(A, 1) + diagIdx - 1, LBound(A, 2) + diagIdx - 1))
        End If
    Next diagIdx

    For j = 1 To n
        ' 非对角元素: L(j, k) = (A(j,k) - Σ L(j,i)·L(k,i)) / L(k,k)
        For k = 1 To j - 1
            s = A(LBound(A, 1) + j - 1, LBound(A, 2) + k - 1)
            For i = 1 To k - 1
                s = s - L(j, i) * L(k, i)
            Next i
            L(j, k) = s / L(k, k)
        Next k

        ' 对角元素: L(j,j) = √(A(j,j) - Σ L(j,i)²)
        s = A(LBound(A, 1) + j - 1, LBound(A, 2) + j - 1)
        For i = 1 To j - 1
            s = s - L(j, i) * L(j, i)
        Next i

        ' 正定性检查: 使用最大对角元作为相对容差 (比单列对角元更稳定)
        If s <= NUM_EPSILON * maxDiag Then
            Err.Raise ERR_NOT_POS_DEF, "CholeskyDecomposition", _
                "矩阵非正定: 对角元素 s[" & j & "]=" & s & " <= 0。"
        End If
        L(j, j) = Sqr(s)
        If L(j, j) > maxDiag Then maxDiag = L(j, j)
    Next j
End Sub

'=============================================================================
' MatrixDeterminant — 行列式 (通过 LU 分解)
'=============================================================================
Public Function MatrixDeterminant(ByRef A As Variant) As Double
    Dim matA() As Double: matA = ToDoubleMatrix(A)
    Dim L() As Double, U() As Double, P() As Double
    Dim n As Long, i As Long, det As Double
    Dim swapCount As Long

    On Error Resume Next
    LUDecomposition matA, L, U, P, swapCount
    If Err.Number = ERR_SINGULAR Then
        ' 奇异矩阵 — 行列式为 0
        On Error GoTo 0
        MatrixDeterminant = 0#: Exit Function
    ElseIf Err.Number <> 0 Then
        ' 非预期错误 — 重新抛出，避免 U 未初始化导致垃圾值
        Dim detErrNum As Long, detErrSrc As String, detErrDesc As String
        detErrNum = Err.Number: detErrSrc = Err.Source: detErrDesc = Err.Description
        On Error GoTo 0
        Err.Raise detErrNum, detErrSrc, detErrDesc
    End If
    On Error GoTo 0
    n = MatrixRows(U)

    det = 1#
    If swapCount Mod 2 = 1 Then det = -1#

    For i = 1 To n
        det = det * U(i, i)
    Next i
    MatrixDeterminant = det
End Function

'=============================================================================
' ForwardSubstitution — 解 Lx = b (L 下三角, 单位对角线可选)
'=============================================================================
Private Function ForwardSub(ByRef L() As Double, ByRef b() As Double, _
                            ByVal n As Long) As Double()
    Dim x() As Double: ReDim x(1 To n)
    Dim i As Long, j As Long, s As Double
    For i = 1 To n
        s = b(i, 1)
        For j = 1 To i - 1
            s = s - L(i, j) * x(j)
        Next j
        x(i) = s / L(i, i)
    Next i
    ForwardSub = x
End Function

'=============================================================================
' BackSubstitution — 解 Ux = b (U 上三角)
'=============================================================================
Private Function BackSub(ByRef U() As Double, ByRef b() As Double, _
                         ByVal n As Long) As Double()
    Dim x() As Double: ReDim x(1 To n)
    Dim i As Long, j As Long, s As Double
    For i = n To 1 Step -1
        s = b(i, 1)
        For j = i + 1 To n
            s = s - U(i, j) * x(j)
        Next j
        x(i) = s / U(i, i)
    Next i
    BackSub = x
End Function

'=============================================================================
' SolveLinearSystem — 解 Ax = b, 自适应路径选择
'
' 方阵 (m = n): LU 部分主元 → 前代/回代 (O(n³/3), 最快)
'   若 LU 检测奇异 → 回退 SVD 伪逆
' 超定 (m > n): QR 经济分解 → 回代 (O(mn² - n³/3))
'   若 R 对角线过小 (秩亏) → 回退 SVD 伪逆
' 欠定 (m < n): SVD 伪逆 → 最小范数解
'=============================================================================
Public Function SolveLinearSystem(ByRef A As Variant, ByRef b As Variant, _
                              Optional ByVal tolerance As Double = -1#) As Double()
    Dim m As Long, n As Long
    ValidateMatrix A, "A"
    m = MatrixRows(A): n = MatrixCols(A)

    ' 检测并处理 b 的维度 (1D 或 2D) (#37)
    Dim bIs1D As Boolean: bIs1D = True
    Err.Clear: On Error Resume Next
    Dim bProbe As Long: bProbe = UBound(b, 2)
    bIs1D = (Err.Number <> 0)
    On Error GoTo 0

    Dim lbB As Long, ubB As Long
    If bIs1D Then
        lbB = LBound(b): ubB = UBound(b)
    Else
        lbB = LBound(b, 1): ubB = UBound(b, 1)
    End If
    If ubB - lbB + 1 <> m Then
        Err.Raise ERR_SOLVE_MISMATCH, "SolveLinearSystem", _
            "b 的维度与 A 的行数不匹配: |b|=" & (ubB - lbB + 1) & ", rows(A)=" & m
    End If

    ' 将 b 转为 (m x 1) 的 2D 矩阵
    Dim b2D() As Double, i As Long
    ReDim b2D(1 To m, 1 To 1)
    For i = 1 To m
        If bIs1D Then
            b2D(i, 1) = b(lbB + i - 1)
        Else
            b2D(i, 1) = b(lbB + i - 1, LBound(b, 2))
        End If
    Next i

    Dim result() As Double

    ' --- Path 1: 方阵 → LU 部分主元 ---
    If m = n Then
        Dim Ad() As Double: Ad = ToDoubleMatrix(A)
        Dim L() As Double, U() As Double, P() As Double, swp As Long
        On Error GoTo LU_SINGULAR
        LUDecomposition Ad, L, U, P, swp
        On Error GoTo 0
        ' Pb
        Dim Pb() As Double: ReDim Pb(1 To n, 1 To 1)
        Dim j As Long
        For i = 1 To n
            Pb(i, 1) = 0#
            For j = 1 To n
                Pb(i, 1) = Pb(i, 1) + P(i, j) * b2D(j, 1)
            Next j
        Next i
        ' Ly = Pb → Ux = y
        Dim y() As Double: y = ForwardSub(L, Pb, n)
        Dim y2D() As Double: ReDim y2D(1 To n, 1 To 1)
        For i = 1 To n: y2D(i, 1) = y(i): Next i
        result = BackSub(U, y2D, n)
        SolveLinearSystem = result
        Exit Function
LU_SINGULAR:
        On Error GoTo 0
        ' 回退 SVD
    End If

    ' --- Path 2: 超定 → QR 经济分解 + 回代 ---
    If m > n Then
        Dim Aqr() As Double: Aqr = ToDoubleMatrix(A)
        Dim Q() As Double, R() As Double
        QRDecomposition Aqr, Q, R, True  ' economy: Q(m×n), R(n×n)
        ' 检查 R 对角线是否秩亏
        Dim rMax As Double: rMax = 0#
        For i = 1 To n
            If Abs(R(i, i)) > rMax Then rMax = Abs(R(i, i))
        Next i
        Dim rankTol As Double
        If tolerance > 0# Then
            rankTol = tolerance
        Else
            rankTol = rMax * NUM_EPSILON * CDbl(m)
        End If
        Dim fullRank As Boolean: fullRank = True
        For i = 1 To n
            If Abs(R(i, i)) < rankTol Then fullRank = False: Exit For
        Next i
        If fullRank Then
            ' Qᵀb
            Dim QtB() As Double: ReDim QtB(1 To n, 1 To 1)
            Dim k As Long
            For k = 1 To n
                QtB(k, 1) = 0#
                For i = 1 To m
                    QtB(k, 1) = QtB(k, 1) + Q(i, k) * b2D(i, 1)
                Next i
            Next k
            result = BackSub(R, QtB, n)
            SolveLinearSystem = result
            Exit Function
        End If
        ' 秩亏 → 回退 SVD
    End If

    ' --- Path 3: SVD 伪逆 (欠定/秩亏回退) ---
    Dim Ap() As Double
    Ap = PseudoInverse(A, tolerance)
    Dim X() As Double
    X = MatrixMultiply(Ap, b2D)
    Dim nOut As Long: nOut = MatrixRows(Ap)
    ReDim result(1 To nOut)
    For i = 1 To nOut
        result(i) = X(i, 1)
    Next i
    SolveLinearSystem = result
End Function

'=============================================================================
' MatrixConditionNumber — 条件数 = max(σ) / min(σ)
'
' 条件数大 → 病态矩阵，微小的输入变化会导致巨大输出变化
'=============================================================================
Public Function MatrixConditionNumber(ByRef A As Variant, _
                                      Optional ByVal tol As Double = DEFAULT_TOL, _
                                      Optional ByVal maxSweeps As Variant = MAX_SWEEPS) As Double

    Dim matA() As Double: matA = ToDoubleMatrix(A)
    Dim U() As Double, S() As Double, Vt() As Double
    Dim i As Long, k As Long
    Dim maxSigma As Double, minSigma As Double

    SVD matA, U, S, Vt, tol, maxSweeps
    k = MatrixRows(S)

    maxSigma = S(1, 1)
    minSigma = S(1, 1)
    For i = 2 To k
        If S(i, i) > maxSigma Then maxSigma = S(i, i)
        If S(i, i) < minSigma Then minSigma = S(i, i)
    Next i

    If maxSigma = 0# Or minSigma < tol * maxSigma Then
        MatrixConditionNumber = MAX_DOUBLE
    Else
        MatrixConditionNumber = maxSigma / minSigma
    End If
End Function

'=============================================================================
' 向量运算
'=============================================================================

' VectorDot — 向量点积 a·b
Public Function VectorDot(ByRef a As Variant, ByRef b As Variant) As Double
    Dim nA As Long, nB As Long
    nA = UBound(a) - LBound(a) + 1
    nB = UBound(b) - LBound(b) + 1
    If nA <> nB Then
        Err.Raise ERR_VEC_DIM, "VectorDot", "向量维度不匹配: |a|=" & nA & ", |b|=" & nB
    End If

    Dim i As Long, result As Double
    Dim la As Long: la = LBound(a)
    Dim lb As Long: lb = LBound(b)
    Dim c As Double, y As Double, t As Double
    result = 0#: c = 0#
    For i = 0 To nA - 1
        y = a(la + i) * b(lb + i) - c: t = result + y: c = (t - result) - y: result = t
    Next i
    VectorDot = result
End Function

' VectorNorm — 向量范数
' normType: "2" = 欧几里得 (默认), "1" = L1 (曼哈顿), "inf" = L_SYM (切比雪夫)
Public Function VectorNorm(ByRef v As Variant, Optional ByVal normType As String = "2") As Double
    Dim n As Long, i As Long
    n = UBound(v) - LBound(v) + 1
    If n < 1 Then
        VectorNorm = 0#
        Exit Function
    End If

    Dim lb As Long: lb = LBound(v)

    Select Case LCase$(normType)
        Case "1"
            ' L1 范数
            Dim sumAbs As Double: sumAbs = 0#
            For i = 0 To n - 1
                sumAbs = sumAbs + Abs(v(lb + i))
            Next i
            VectorNorm = sumAbs
        Case "inf", "infinity"
            ' L_SYM 范数
            Dim maxAbs As Double: maxAbs = 0#
            For i = 0 To n - 1
                If Abs(v(lb + i)) > maxAbs Then maxAbs = Abs(v(lb + i))
            Next i
            VectorNorm = maxAbs
        Case Else
            ' L2 范数 (scaled to prevent overflow)
            Dim maxVal As Double: maxVal = 0#
            For i = 0 To n - 1
                If Abs(v(lb + i)) > maxVal Then maxVal = Abs(v(lb + i))
            Next i
            If maxVal = 0# Then VectorNorm = 0#: Exit Function
            ' Scale and accumulate
            Dim ss As Double: ss = 0#
            For i = 0 To n - 1
                ss = ss + (v(lb + i) / maxVal) ^ 2
            Next i
            VectorNorm = maxVal * Sqr(ss)
    End Select
End Function

' VectorCross — 三维向量叉积 a × b
Public Function VectorCross(ByRef a As Variant, ByRef b As Variant) As Double()
    Dim nA As Long, nB As Long
    nA = UBound(a) - LBound(a) + 1
    nB = UBound(b) - LBound(b) + 1
    If nA <> 3 Or nB <> 3 Then
        Err.Raise ERR_VEC_3D_ONLY, "VectorCross", "叉积仅支持三维向量: |a|=" & nA & ", |b|=" & nB
    End If

    Dim result() As Double
    Dim la As Long: la = LBound(a)
    Dim lb As Long: lb = LBound(b)

    ReDim result(1 To 3)
    result(1) = a(la + 1) * b(lb + 2) - a(la + 2) * b(lb + 1)
    result(2) = a(la + 2) * b(lb + 0) - a(la + 0) * b(lb + 2)
    result(3) = a(la + 0) * b(lb + 1) - a(la + 1) * b(lb + 0)
    VectorCross = result
End Function

'=============================================================================
' 工作表 UDF 扩展
'=============================================================================

Public Function UDF_LINALG_DET(ByVal rng As Variant) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double
    A = RangeToMatrix(rng)
    UDF_LINALG_DET = MatrixDeterminant(A)
    Exit Function
ErrHandler:
    UDF_LINALG_DET = CVErr(xlErrValue)
End Function

' Private helper — build Double() vector from Range, array, or scalar input.
' Err.Raise is allowed here (not a UDF — called from UDF wrapper only).
Private Function BuildSolveVector(ByVal src As Variant) As Double()
    Dim v As Variant, n As Long, i As Long, j As Long
    Dim result() As Double

    If IsObject(src) Then
        If TypeOf src Is Range Then
            v = src.Value
            If Not IsArray(v) Then
                ' Single cell → scalar
                If Not IsNumeric(v) Then Err.Raise ERR_INVALID_INPUT, "BuildSolveVector", "Cell is non-numeric"
                ReDim result(1 To 1): result(1) = CDbl(v)
                BuildSolveVector = result: Exit Function
            End If
            If src.Rows.Count = 1 Then
                n = src.Columns.Count: ReDim result(1 To n)
                For j = 1 To n
                    If Not IsNumeric(v(1, j)) Then Err.Raise ERR_INVALID_INPUT, "BuildSolveVector", "Element (" & j & ") non-numeric"
                    result(j) = CDbl(v(1, j))
                Next j
            Else
                n = src.Rows.Count: ReDim result(1 To n)
                For i = 1 To n
                    If Not IsNumeric(v(i, 1)) Then Err.Raise ERR_INVALID_INPUT, "BuildSolveVector", "Element (" & i & ") non-numeric"
                    result(i) = CDbl(v(i, 1))
                Next i
            End If
            BuildSolveVector = result: Exit Function
        End If
    End If

    ' Array or scalar (no Range object)
    If IsArray(src) Then
        v = src
    ElseIf IsNumeric(src) Then
        ReDim result(1 To 1): result(1) = CDbl(src)
        BuildSolveVector = result: Exit Function
    Else
        Err.Raise ERR_INVALID_INPUT, "BuildSolveVector", "Input is not numeric"
    End If

    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    Dim is2D As Boolean
    lb1 = LBound(v): ub1 = UBound(v)
    Err.Clear
    On Error Resume Next: lb2 = LBound(v, 2): ub2 = UBound(v, 2): is2D = (Err.Number = 0): On Error GoTo 0
    If Not is2D Then
        ' 1D array
        n = ub1 - lb1 + 1: ReDim result(1 To n)
        For i = 1 To n
            If Not IsNumeric(v(lb1 + i - 1)) Then Err.Raise ERR_INVALID_INPUT, "BuildSolveVector", "Element non-numeric"
            result(i) = CDbl(v(lb1 + i - 1))
        Next i
    Else
        ' 2D — take first column or first row
        If ub1 - lb1 + 1 >= ub2 - lb2 + 1 Then
            n = ub1 - lb1 + 1: ReDim result(1 To n)
            For i = 1 To n
                If Not IsNumeric(v(lb1 + i - 1, lb2)) Then Err.Raise ERR_INVALID_INPUT, "BuildSolveVector", "Element non-numeric"
                result(i) = CDbl(v(lb1 + i - 1, lb2))
            Next i
        Else
            n = ub2 - lb2 + 1: ReDim result(1 To n)
            For j = 1 To n
                If Not IsNumeric(v(lb1, lb2 + j - 1)) Then Err.Raise ERR_INVALID_INPUT, "BuildSolveVector", "Element non-numeric"
                result(j) = CDbl(v(lb1, lb2 + j - 1))
            Next j
        End If
    End If
    BuildSolveVector = result
End Function

Public Function UDF_LINALG_SOLVE(ByVal rngA As Variant, ByVal rngB As Variant) As Variant
    On Error GoTo EH
    Dim A() As Double: A = RangeToMatrix(rngA)
    Dim b() As Double: b = BuildSolveVector(rngB)
    UDF_LINALG_SOLVE = DblMatrixToVariant(ArrayToColumn(SolveLinearSystem(A, b)))
    Exit Function
EH: UDF_LINALG_SOLVE = CVErr(xlErrValue)
End Function

Public Function UDF_LINALG_CHOLESKY(ByVal rng As Variant) As Variant
    On Error GoTo ErrHandler
    Dim A() As Double, L() As Double
    A = RangeToMatrix(rng)
    CholeskyDecomposition A, L
    UDF_LINALG_CHOLESKY = DblMatrixToVariant(L)
    Exit Function
ErrHandler:
    UDF_LINALG_CHOLESKY = CVErr(xlErrValue)
End Function


Public Function UDF_LINALG_POLYFIT(ByVal rngX As Variant, ByVal rngY As Variant, Optional ByVal degree As Variant = 1) As Variant
    On Error GoTo EH: UDF_LINALG_POLYFIT = PolyFit(rngX, rngY, degree): Exit Function
EH: UDF_LINALG_POLYFIT = CVErr(xlErrValue)
End Function
'=============================================================================
' PolyFit — 最小二乘多项式拟合 (返回系数，高次到低次)
'
' 使用 QR 分解求解最小二乘问题 min ||X*beta - y||，避免正规方程 (X^T X)^(-1) X^T y
' 在 Vandermonde 矩阵上平方条件数的问题。依赖 QRDecomposition + MatrixMultiply。
' 置于本模块而非 RegressUtils。UDF 名遵循 LINALG 前缀以保持命名空间一致。
'
' degree=1→[a,b] for y=ax+b; degree=2→[a,b,c] for y=ax²+bx+c
'=============================================================================
Public Function PolyFit(ByRef rngX As Variant, ByRef rngY As Variant, _
    Optional ByVal degree As Long = 1) As Variant
    ' 提取输入为 Double 数组
    Dim x() As Double, y() As Double
    On Error GoTo ErrHandler
    Dim rng As Range
    If IsObject(rngX) Then
        If TypeOf rngX Is Range Then
            Set rng = rngX
            x = RangeToMatrix(rng)
        End If
    ElseIf IsArray(rngX) Then
        x = VariantToDouble1D(rngX)
    End If
    If IsObject(rngY) Then
        If TypeOf rngY Is Range Then
            Set rng = rngY
            y = RangeToMatrix(rng)
        End If
    ElseIf IsArray(rngY) Then
        y = VariantToDouble1D(rngY)
    End If
    ' Normalize: RangeToMatrix returns 2D column (1 To n, 1 To 1);
    ' VariantToDouble1D returns 1D. Unify to 1D for single-subscript access.
    Dim xi As Long, x2d As Boolean, y2d As Boolean
    On Error Resume Next
    xi = UBound(x, 2): x2d = (Err.Number = 0): Err.Clear
    xi = UBound(y, 2): y2d = (Err.Number = 0): Err.Clear
    On Error GoTo ErrHandler
    If x2d Then
        Dim xTmp() As Double: ReDim xTmp(LBound(x, 1) To UBound(x, 1))
        For xi = LBound(x, 1) To UBound(x, 1): xTmp(xi) = x(xi, 1): Next xi
        x = xTmp
    End If
    If y2d Then
        Dim yt() As Double: ReDim yt(LBound(y, 1) To UBound(y, 1))
        For xi = LBound(y, 1) To UBound(y, 1): yt(xi) = y(xi, 1): Next xi
        y = yt
    End If
    Dim nX As Long: nX = UBound(x) - LBound(x) + 1
    Dim nY As Long: nY = UBound(y) - LBound(y) + 1
    If nX <> nY Then Err.Raise ERR_INVALID_SIZE, "PolyFit", "x 与 y 长度不匹配: " & nX & " vs " & nY
    If nX < degree + 1 Then Err.Raise ERR_INVALID_SIZE, "PolyFit", _
        "数据点不足: 需要至少 " & (degree + 1) & " 个点，实际只有 " & nX & " 个。"
    Dim n As Long: n = nX
    Dim lb As Long: lb = LBound(x)
    ' 构建设计矩阵
    Dim Xmat() As Double, i As Long, j As Long
    ReDim Xmat(1 To n, 1 To degree + 1)
    For i = 1 To n
        Xmat(i, 1) = 1#
        For j = 2 To degree + 1
            Xmat(i, j) = x(lb + i - 1) ^ (j - 1)
        Next j
    Next i
    ' QR 最小二乘: 分解 X = Q*R (经济模式), 解 R*beta = Q^T*y
    ' 避免正规方程 (X^T X)^(-1) X^T y 平方条件数的问题
    Dim Qmat() As Double, Rmat() As Double
    QRDecomposition Xmat, Qmat, Rmat, True  ' economy=True: Q(mxk), R(kxk)
    Dim yCol() As Double: ReDim yCol(1 To n, 1 To 1)
    For i = 1 To n: yCol(i, 1) = y(lb + i - 1): Next i
    ' 计算 Q^T * y
    Dim QtY() As Double: QtY = MatrixMultiply(MatrixTranspose(Qmat), yCol)
    ' 回代求解 R * beta = QtY (R 是 k×k 上三角)
    Dim k As Long: k = degree + 1
    Dim coeffs() As Double: ReDim coeffs(1 To k)
    Dim row As Long, col As Long, s As Double
    For row = k To 1 Step -1
        s = QtY(row, 1)
        For col = row + 1 To k
            s = s - Rmat(row, col) * coeffs(col)
        Next col
        ' 标度感知容差以 R(1,1) 为参考 (QR 分解中通常最大)
        If Abs(Rmat(row, row)) < 1E-14 * (1# + Abs(Rmat(1, 1))) Then
            coeffs(row) = 0#
        Else
            coeffs(row) = s / Rmat(row, row)
        End If
    Next row
    ' 返回系数（高次→低次）
    Dim result() As Variant: ReDim result(1 To degree + 1, 1 To 1)
    For i = 1 To degree + 1: result(i, 1) = coeffs(degree + 2 - i): Next i
    PolyFit = result
    Exit Function
ErrHandler:
    ' 重新抛出未预期错误 (PolyFit 不是 UDF, 不应返回 CVErr)
    Err.Raise Err.Number, "PolyFit", Err.Description
End Function

' 辅助：Variant 1D 数组 → Double 1D 数组
' 2D 数组（如 Range.Value）自动提取第一列
Private Function VariantToDouble1D(ByRef arr As Variant) As Double()
    Dim result() As Double, i As Long, v As Variant
    Dim lb As Long, ub As Long

    ' 检测数组维度 — 2D 数组提取第一列, 避免静默丢弃其余列
    On Error Resume Next
    ub = UBound(arr, 2)
    If Err.Number = 0 Then
        ' 2D 数组: 显式提取第一列 (Range.Value 典型为 (1 to n, 1 to 1))
        Err.Clear: On Error GoTo 0
        lb = LBound(arr, 1): ub = UBound(arr, 1)
        ReDim result(lb To ub)
        Dim colLB As Long: colLB = LBound(arr, 2)
        For i = lb To ub
            v = arr(i, colLB)
            If VarType(v) <> vbBoolean And IsNumeric(v) Then result(i) = CDbl(v) Else result(i) = 0#
        Next i
    Else
        ' 1D 数组
        Err.Clear: On Error GoTo 0
        lb = LBound(arr): ub = UBound(arr)
        ReDim result(lb To ub)
        For i = lb To ub
            v = arr(i)
            If VarType(v) <> vbBoolean And IsNumeric(v) Then result(i) = CDbl(v) Else result(i) = 0#
        Next i
    End If
    VariantToDouble1D = result
End Function

Private Function ArrayToColumn(ByRef arr() As Double) As Double()
    Dim result() As Double
    Dim n As Long, i As Long
    n = UBound(arr) - LBound(arr) + 1
    ReDim result(1 To n, 1 To 1)
    For i = 1 To n
        result(i, 1) = arr(LBound(arr) + i - 1)
    Next i
    ArrayToColumn = result
End Function

'=====================================================================
' 使用示例
'=====================================================================
' 1. VBA 中调用:
'   Dim A() As Double, U() As Double, S() As Double, Vt() As Double
'   A = RangeToMatrix(ThisWorkbook.Sheets("Sheet1").Range("A1:C3"))
'   SVD A, U, S, Vt
'   MatrixToRange U, ThisWorkbook.Sheets("Sheet1").Range("E1")
'
'   MatrixTrace(A)                                        → 迹
'   MatrixDeterminant(A)                                  → 行列式
'   MatrixConditionNumber(A)                              → 条件数
'   Dim C() As Double: C = MatrixAdd(A, B)                  ' 加法
'   Dim x() As Double: x = SolveLinearSystem(A, b)          ' 解 Ax=b
'   Dim L() As Double, U2() As Double, P() As Double
'   LUDecomposition A, L, U2, P                             ' LU 分解
'   Dim Lchol() As Double
'   CholeskyDecomposition A, Lchol                          ' Cholesky 分解
'   VectorDot(v1, v2)                                     → 点积
'
' 2. 工作表数组公式:
'   =SVD_U(A1:C3)        (选中输出区域, Ctrl+Shift+Enter)
'   =EIGVAL_SYM(D1:F3)
'   =PINV(G1:H4)
'   =MATRIXRANK(J1:L5, 1E-10)
'   =MATRIXDET(A1:C3)
'   =MATRIXSOLVE(A1:C3, D1:D3)
'   =CHOLESKY_L(A1:C3)    (对称正定矩阵 → 下三角 L)

'=====================================================================