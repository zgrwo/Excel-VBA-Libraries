Option Explicit

'==============================================================================
' Module:       RegressUtils
' Purpose:      Regression: OLS, WLS, Ridge, ANOVA, factor importance
' Layer:        Statistics
' Dependencies: VBA-Core + LinearUtils + StatsUtils (import these first)
' Public:       16 functions/subs
' Notes:        Requires LinearUtils and StatsUtils loaded before this module.
'==============================================================================

Private DP As New DictProxy
Private VK As New VariantKit

'=====================================================================
' RegressUtils.bas — 回归分析与统计建模
'
' 公共函数:

'   FactorImportance     — 因子重要性排序 (标准化回归系数)
'   InteractionEffects   — 因子交互效应检测
'   ANOVAOneWay          — 单因素方差分析 (返回 Dictionary)
'   ANOVAOneWay_Fstat    — 单因素方差分析 (仅 F 统计量)
'   LinearModelFit       — 多元线性回归拟合 (返回 Dictionary)
'   LinearModelPredict   — 基于模型预测结果
'   FactorSweep          — 单因子扫描分析
'   OptimizeFactors      — 最优因子组合搜索
'
' 工作表函数 (UDF_REGRESS_*):
'   UDF_REGRESS_CORREL, UDF_REGRESS_IMPORTANCE, UDF_REGRESS_INTERACT, UDF_REGRESS_ANOVA, UDF_REGRESS_PREDICT, UDF_REGRESS_SWEEP, UDF_REGRESS_OPTIMIZE
'
' 依赖:
'   - LinearUtils.bas (MatrixMultiply, MatrixTranspose, PseudoInverse, QRDecomposition)
'   - StatsUtils.bas (TDist2T, FDistRT — 用于 p 值计算; CorrelationMatrix — 用于 UDF_REGRESS_CORREL)
'=====================================================================


' Numerical tolerance constants
Private Const NUM_EPS    As Double = 2.22044604925031E-16  ' Machine epsilon
Private Const RANK_TOL As Double = 1E-15                 ' Rank determination tolerance

Private Const MAX_CATEGORICAL_LEVELS As Long = 50
Private Const MAX_GRID_COMBOS As Long = 200000
Private Const PI As Double = 3.14159265358979
Private Const ERR_INVALID_DATA  As Long = vbObjectError + 3001
Private Const ERR_UNDERDETERM   As Long = vbObjectError + 3002
Private Const ERR_UNKNOWN_LEVEL As Long = vbObjectError + 3004
Private Const ERR_TOO_FEW_ROWS  As Long = vbObjectError + 3101
Private Const REG_IDX_THRESHOLD As Long = 16


'=============================================================================
' 类型检测辅助函数
'=============================================================================

Private Function IsNumericCell(ByVal v As Variant) As Boolean
    IsNumericCell = VK.IsNumericCell(v)
End Function

Private Function ToDouble(ByVal v As Variant) As Double
    If IsNull(v) Or IsEmpty(v) Then
        ToDouble = 0#
    ElseIf IsError(v) Then
        Err.Raise ERR_INVALID_DATA, "ToDouble", "Error value in data — check for #N/A or #DIV/0! cells"
    ElseIf VarType(v) = vbBoolean Then
        If CBool(v) Then ToDouble = 1# Else ToDouble = 0#
    ElseIf IsNumeric(v) Then
        ToDouble = CDbl(v)
    Else
        Err.Raise ERR_INVALID_DATA, "ToDouble", "Non-numeric value '" & CStr(v) & "' in data column"
    End If
End Function

Private Function ToBoolDouble(ByVal v As Variant) As Double
    If IsEmpty(v) Then ToBoolDouble = 0#: Exit Function
    If VarType(v) = vbBoolean Then
        If CBool(v) Then ToBoolDouble = 1# Else ToBoolDouble = 0#
    ElseIf IsNumeric(v) Then
        If CDbl(v) <> 0# Then ToBoolDouble = 1# Else ToBoolDouble = 0#
    End If
End Function

'=============================================================================
' DetectColumnType — 检测列的数据类型
'
' 返回: "numeric" / "boolean" / "categorical"
'=============================================================================
Private Function DetectColumnType(ByRef dataArr As Variant, ByVal col As Long, _
                                   ByVal startRow As Long, ByVal endRow As Long) As String
    Dim r As Long, hasNum As Boolean, hasBool As Boolean, hasText As Boolean
    Dim v As Variant
    Dim s As String

    For r = startRow To endRow
        v = dataArr(r, col)
        If Not IsEmpty(v) Then
            If VarType(v) = vbBoolean Then
                hasBool = True
            ElseIf IsNumeric(v) Then
                hasNum = True
            ElseIf VarType(v) = vbString Then
                s = Trim(CStr(v))
                If Len(s) > 0 Then hasText = True
            Else
                hasText = True
            End If
        End If
    Next

    If hasText Then
        DetectColumnType = "categorical"
    ElseIf hasBool Then
        DetectColumnType = "boolean"
    ElseIf hasNum Then
        ' 如果只有 0/1 且至少有一个真正的 Boolean，视为 boolean
        DetectColumnType = "numeric"
    Else
        DetectColumnType = "categorical"
    End If
End Function

'=============================================================================
' GetCategoricalLevels — 获取分类列的所有不同值
'=============================================================================
Private Function GetCategoricalLevels(ByRef dataArr As Variant, ByVal col As Long, _
                                       ByVal startRow As Long, ByVal endRow As Long) As String()
    Dim dict As Object: Set dict = DP.Create()
    Dim r As Long, v As Variant, key As Variant
    Dim result() As String
    Dim i As Long
    Dim j As Long, nLev As Long
    Dim tmp As String

    For r = startRow To endRow
        v = dataArr(r, col)
        If Not IsEmpty(v) Then
            key = CStr(v)
            If Not dict.Exists(key) Then dict.Add key, key
        End If
    Next

    If dict.Count = 0 Then
        ' 返回类型安全的空 String 数组（替代 Array() 返回 Variant）
        Dim emptyStr() As String
        GetCategoricalLevels = emptyStr
        Exit Function
    End If
    ReDim result(0 To dict.Count - 1)
    For Each key In dict.Keys
        result(i) = key
        i = i + 1
    Next

    ' 对水平排序，确保训练/预测时列顺序一致
    nLev = UBound(result) - LBound(result) + 1
    If nLev > 1 Then
        For i = LBound(result) To UBound(result) - 1
            For j = i + 1 To UBound(result)
                If StrComp(result(i), result(j), vbTextCompare) > 0 Then
                    tmp = result(i): result(i) = result(j): result(j) = tmp
                End If
            Next
        Next
    End If

    GetCategoricalLevels = result
End Function

'=============================================================================
' BuildDesignMatrix — 构建设计矩阵
'
' X 的列结构: [intercept (1s), numeric_cols..., boolean_cols..., dummy_cols...]
' 每个分类因子 (k 个水平) 产生 k-1 个哑变量列 (丢弃第一个水平作为参考)
'
' 返回 Dictionary:
'   "X"               — Double(1..n, 1..p)  设计矩阵
'   "y"               — Double(1..n)         目标向量
'   "coef_names"      — String(1..p)         系数名称
'   "factor_map"      — Dictionary           编码映射 (用于预测)
'=============================================================================
Private Function BuildDesignMatrix(ByRef dataArr As Variant, _
                                    ByVal factorCols As Variant, _
                                    ByVal resultCol As Long, _
                                    ByVal firstDataRow As Long, _
                                    ByVal lastDataRow As Long, _
                                    Optional ByVal hasHeader As Boolean = True) As Object
    Dim n As Long: n = lastDataRow - firstDataRow + 1
    Dim numFactors As Long
    Dim i As Long, j As Long, r As Long, c As Long
    Dim p As Long
    Dim colTypes() As String
    Dim colIdx() As Long
    Dim catLevels() As String
    Dim fi As Long, fCol As Long
    Dim colInfo As Object
    Dim cf As Long
    Dim nLevels As Long
    Dim validRow() As Boolean
    Dim validCount As Long
    Dim isRowValid As Boolean
    Dim X() As Double
    Dim y() As Double
    Dim coefNames() As String
    Dim designCol As Long
    Dim factorMap As Object
    Dim colName As String
    Dim numSum As Double, numValid As Long, nv As Double
    Dim boolSum As Double, bv As Double, boolValid As Long
    Dim dummyCols() As Long
    Dim Xc() As Double, yc() As Double
    Dim src As Long, dst As Long
    Dim result As Object

    ' 确定因子维度 — 标量自动包装为单元素数组
    If Not IsArray(factorCols) Then
        factorCols = Array(factorCols)
    End If
    numFactors = UBound(factorCols) - LBound(factorCols) + 1

    ' 遍历因子，计算设计矩阵列数
    p = 1  ' 截距
    ReDim colTypes(1 To numFactors)
    ReDim colIdx(1 To numFactors)

    For fi = 1 To numFactors
        fCol = CLng(factorCols(LBound(factorCols) + fi - 1))
        colIdx(fi) = fCol
        colTypes(fi) = DetectColumnType(dataArr, fCol, firstDataRow, lastDataRow)

        Select Case colTypes(fi)
            Case "numeric", "boolean"
                p = p + 1
            Case "categorical"
                catLevels = GetCategoricalLevels(dataArr, fCol, firstDataRow, lastDataRow)
                If UBound(catLevels) - LBound(catLevels) + 1 > MAX_CATEGORICAL_LEVELS Then
                    Err.Raise ERR_INVALID_DATA, "BuildDesignMatrix", _
                        "列 " & fCol & " 的分类水平数 (" & (UBound(catLevels) - LBound(catLevels) + 1) & ") 超过上限 (" & MAX_CATEGORICAL_LEVELS & ")"
                End If
                nLevels = UBound(catLevels) - LBound(catLevels) + 1
                If nLevels > 1 Then p = p + nLevels - 1  ' k-1 个哑变量
        End Select
    Next

    ' 列表删除: 识别有效行 (所有因子列和结果列均无缺失值)
    ReDim validRow(1 To n)
    validCount = 0
    For r = 1 To n
        isRowValid = True
        ' 检查结果列: 必须是数值
        If Not IsNumericCell(dataArr(firstDataRow + r - 1, resultCol)) Then
            isRowValid = False
        End If
        ' 检查所有因子列
        If isRowValid Then
            For fi = 1 To numFactors
                fCol = CLng(factorCols(LBound(factorCols) + fi - 1))
                Select Case colTypes(fi)
                    Case "numeric"
                        If Not IsNumericCell(dataArr(firstDataRow + r - 1, fCol)) Then
                            isRowValid = False: Exit For
                        End If
                    Case "boolean"
                        If IsEmpty(dataArr(firstDataRow + r - 1, fCol)) Then
                            isRowValid = False: Exit For
                        End If
                        If IsNull(dataArr(firstDataRow + r - 1, fCol)) Then
                            isRowValid = False: Exit For
                        End If
                        If IsError(dataArr(firstDataRow + r - 1, fCol)) Then
                            isRowValid = False: Exit For
                        End If
                    ' categorical: 总是有效 (空字符串也是有效类别)
                End Select
            Next
        End If
        validRow(r) = isRowValid
        If isRowValid Then validCount = validCount + 1
    Next

    If validCount < 2 Then
        Err.Raise ERR_INVALID_DATA, "BuildDesignMatrix", _
            "有效数据行不足 (需要 >=2 行, 实际 " & validCount & " 行)。请检查缺失值。"
    End If

    ' 分配设计矩阵和 y 向量 (按总行数分配, 稍后压缩)
    ReDim X(1 To n, 1 To p)
    ReDim y(1 To n)
    ReDim coefNames(1 To p)
    coefNames(1) = "(Intercept)"

    ' 填充截距列
    For r = 1 To n
        X(r, 1) = 1#
    Next

    ' 填充因子列并记录映射
    designCol = 1  ' 已填充到截距
    Set factorMap = DP.Create()
    colName = ""

    For fi = 1 To numFactors
        fCol = colIdx(fi)
        If hasHeader Then
            colName = CStr(dataArr(firstDataRow - 1, fCol))
        Else
            colName = "F" & fi
        End If

        Set colInfo = DP.Create()
        colInfo.Add "type", colTypes(fi)
        colInfo.Add "original_col", fCol

        Select Case colTypes(fi)
            Case "numeric"
                designCol = designCol + 1
                colInfo.Add "design_cols", Array(designCol)
                coefNames(designCol) = colName
                numSum = 0#
                numValid = 0
                For r = 1 To n
                    If validRow(r) Then
                        nv = CDbl(dataArr(firstDataRow + r - 1, fCol))
                        X(r, designCol) = nv
                        numSum = numSum + nv
                        numValid = numValid + 1
                    End If
                Next
                If numValid > 0 Then
                    colInfo.Add "mean", numSum / numValid
                Else
                    colInfo.Add "mean", 0#
                End If

            Case "boolean"
                designCol = designCol + 1
                colInfo.Add "design_cols", Array(designCol)
                coefNames(designCol) = colName
                boolSum = 0#
                boolValid = 0
                For r = 1 To n
                    If validRow(r) Then
                        bv = ToBoolDouble(dataArr(firstDataRow + r - 1, fCol))
                        X(r, designCol) = bv
                        boolSum = boolSum + bv
                        boolValid = boolValid + 1
                    End If
                Next
                If boolValid > 0 Then
                    colInfo.Add "mean", boolSum / boolValid
                Else
                    colInfo.Add "mean", 0#
                End If

            Case "categorical"
                catLevels = GetCategoricalLevels(dataArr, fCol, firstDataRow, lastDataRow)
                nLevels = UBound(catLevels) - LBound(catLevels) + 1
                If nLevels > 1 Then
                    ' 第一个水平作为参考 (丢弃), 其余 k-1 个水平各一列
                    ReDim dummyCols(1 To nLevels - 1)
                    For cf = 2 To nLevels
                        designCol = designCol + 1
                        dummyCols(cf - 1) = designCol
                        coefNames(designCol) = colName & "_" & catLevels(LBound(catLevels) + cf - 1)
                        For r = 1 To n
                            If validRow(r) Then
                                If CStr(dataArr(firstDataRow + r - 1, fCol)) = catLevels(LBound(catLevels) + cf - 1) Then
                                    X(r, designCol) = 1#
                                Else
                                    X(r, designCol) = 0#
                                End If
                            End If
                        Next
                    Next
                    colInfo.Add "design_cols", dummyCols
                    colInfo.Add "levels", catLevels
                Else
                    ' 只有一个水平: 没有哑变量列
                    colInfo.Add "design_cols", Array()
                    colInfo.Add "levels", catLevels
                End If
        End Select

        factorMap.Add CStr(fCol), colInfo
    Next

    ' 填充 y 向量
    For r = 1 To n
        If validRow(r) Then
            y(r) = CDbl(dataArr(firstDataRow + r - 1, resultCol))
        End If
    Next

    ' 列表删除: 压缩矩阵, 移除无效行
    If validCount < n Then
        ReDim Xc(1 To validCount, 1 To p)
        ReDim yc(1 To validCount)
        src = 0
        dst = 0
        For src = 1 To n
            If validRow(src) Then
                dst = dst + 1
                For c = 1 To p
                    Xc(dst, c) = X(src, c)
                Next
                yc(dst) = y(src)
            End If
        Next
        X = Xc
        y = yc
        n = validCount
    End If

    Set result = DP.Create()
    result.Add "X", X
    result.Add "y", y
    result.Add "coef_names", coefNames
    result.Add "factor_map", factorMap
    result.Add "n", n
    result.Add "p", p
    If n <= p Then
        Err.Raise ERR_UNDERDETERM, "BuildDesignMatrix", _
            "观测数 (" & n & ") 小于等于参数个数 (" & p & ")，模型欠定。请减少因子或增加数据。"
    End If
    Set BuildDesignMatrix = result
End Function

'=============================================================================
' FitCoefOnly — 仅计算回归系数 (QR 分解)
'
' 比 FitOLS 轻量, 不计算标准误/t值/p值/拟合值等
'=============================================================================
Private Function FitCoefOnly(ByRef X() As Double, ByRef y() As Double) As Double()
    Dim n As Long, p As Long
    Dim Q() As Double, R() As Double
    Dim Qt() As Double, y2D() As Double, Qty() As Double
    Dim i As Long, j As Long, k As Long
    Dim maxAbsR As Double, rTol As Double
    Dim rr0 As Long, rc0 As Long
    Dim coef() As Double
    Dim s As Double

    n = UBound(X, 1) - LBound(X, 1) + 1
    p = UBound(X, 2) - LBound(X, 2) + 1

    QRDecomposition X, Q, R, True

    ' Qᵀy
    Qt = MatrixTranspose(Q)
    ReDim y2D(1 To n, 1 To 1)
    For i = 1 To n: y2D(i, 1) = y(LBound(y) + i - 1): Next
    Qty = MatrixMultiply(Qt, y2D)

    ' 尺度感知容差
    maxAbsR = 0#
    rr0 = LBound(R, 1): rc0 = LBound(R, 2)
    For j = 1 To p
        For k = j To p
            If Abs(R(rr0 + j - 1, rc0 + k - 1)) > maxAbsR Then maxAbsR = Abs(R(rr0 + j - 1, rc0 + k - 1))
        Next k
    Next
    rTol = maxAbsR * NUM_EPS * CDbl(p)
    If rTol < RANK_TOL Then rTol = RANK_TOL

    ' 回代求解
    ReDim coef(1 To p)
    For j = p To 1 Step -1
        s = Qty(LBound(Qty, 1) + j - 1, LBound(Qty, 2))
        For k = j + 1 To p
            s = s - R(rr0 + j - 1, rc0 + k - 1) * coef(k)
        Next
        If Abs(R(rr0 + j - 1, rc0 + j - 1)) < rTol Then
            coef(j) = 0#
        Else
            coef(j) = s / R(rr0 + j - 1, rc0 + j - 1)
        End If
    Next

    FitCoefOnly = coef
End Function

'=============================================================================
' TriangularInverse — 上三角矩阵求逆
'
' R 是 p×p 上三角矩阵。返回 R⁻¹ (也是上三角)。
' 用于 FitOLS 中从 R 直接计算 (XᵀX)⁻¹ = R⁻¹(R⁻¹)ᵀ，避免显式构造 XᵀX
' 平方条件数。
'=============================================================================
Private Function TriangularInverse(ByRef R() As Double) As Double()
    Dim p As Long, i As Long, j As Long, k As Long
    Dim s As Double
    Dim rr0 As Long, rc0 As Long
    Dim Rinv() As Double
    Dim maxDiag As Double, rTol As Double

    p = UBound(R, 2) - LBound(R, 2) + 1
    rr0 = LBound(R, 1): rc0 = LBound(R, 2)
    ReDim Rinv(rr0 To rr0 + p - 1, rc0 To rc0 + p - 1)

    ' Compute scale-aware tolerance for near-zero diagonal elements
    maxDiag = 0#
    For i = 1 To p
        If Abs(R(rr0 + i - 1, rc0 + i - 1)) > maxDiag Then
            maxDiag = Abs(R(rr0 + i - 1, rc0 + i - 1))
        End If
    Next
    rTol = maxDiag * NUM_EPS * CDbl(p)
    If rTol < RANK_TOL Then rTol = RANK_TOL

    For j = p To 1 Step -1
        If Abs(R(rr0 + j - 1, rc0 + j - 1)) < rTol Then
            Rinv(rr0 + j - 1, rc0 + j - 1) = 0#
        Else
            Rinv(rr0 + j - 1, rc0 + j - 1) = 1# / R(rr0 + j - 1, rc0 + j - 1)
        End If
        For i = j - 1 To 1 Step -1
            s = 0#
            For k = i + 1 To j
                s = s + R(rr0 + i - 1, rc0 + k - 1) * Rinv(rr0 + k - 1, rc0 + j - 1)
            Next
            If Abs(R(rr0 + i - 1, rc0 + i - 1)) < rTol Then
                Rinv(rr0 + i - 1, rc0 + j - 1) = 0#
            Else
                Rinv(rr0 + i - 1, rc0 + j - 1) = -s / R(rr0 + i - 1, rc0 + i - 1)
            End If
        Next
    Next
    TriangularInverse = Rinv
End Function

'=============================================================================
' FitOLS — 普通最小二乘回归
'
' 使用 QR 分解求系数与标准误差。系数通过回代 Rβ = Qᵀy 求解；
' 标准误差的 (XᵀX)⁻¹ 从 R⁻¹(R⁻¹)ᵀ 直接计算，避免显式构造 XᵀX
' （构造 XᵀX 会将条件数平方，导致病态矩阵精度损失）。
'=============================================================================
Public Function FitOLS(ByRef X As Variant, ByRef y As Variant) As Object
    Dim n As Long, p As Long
    Dim Q() As Double, R() As Double
    Dim Qt() As Double, y2D() As Double, Qty() As Double
    Dim i As Long, j As Long, k As Long
    Dim maxAbsR As Double, rTol As Double
    Dim rr0 As Long, rc0 As Long
    Dim coef() As Double
    Dim s As Double
    Dim XtXinv() As Double, Rinv() As Double, RinvT() As Double
    Dim fitted() As Double, resid() As Double
    Dim yMean As Double, sst As Double, sse As Double
    Dim sstC As Double, sseC As Double
    Dim dy As Double, dt As Double
    Dim r2 As Double, adjR2 As Double
    Dim dfReg As Long, dfRes As Long
    Dim sigma2 As Double
    Dim se() As Double
    Dim tStats() As Double, pValues() As Double
    Dim fStat As Double, fPValue As Double
    Dim ssReg As Double
    Dim model As Object

    VK.NormalizeInput X
    VK.NormalizeInput y, True  ' 2D column → 1D

    n = UBound(X, 1) - LBound(X, 1) + 1
    p = UBound(X, 2) - LBound(X, 2) + 1

    ' 提取 Double 数组用于 QRDecomposition（要求 typed Double 参数）
    Dim Xqr() As Double, ri As Long, ci As Long
    Dim rLo As Long: rLo = LBound(X, 1): Dim rHi As Long: rHi = UBound(X, 1)
    Dim cLo As Long: cLo = LBound(X, 2): Dim cHi As Long: cHi = UBound(X, 2)
    ReDim Xqr(rLo To rHi, cLo To cHi)
    For ri = rLo To rHi
        For ci = cLo To cHi
            Xqr(ri, ci) = CDbl(X(ri, ci))
        Next
    Next

    ' QR 分解用于稳定计算 β 系数
    QRDecomposition Xqr, Q, R, True  ' 经济模式: Q 为 m×p, R 为 p×p

    ' Qᵀy
    Qt = MatrixTranspose(Q)
    ReDim y2D(1 To n, 1 To 1)
    For i = 1 To n: y2D(i, 1) = y(LBound(y) + i - 1): Next
    Qty = MatrixMultiply(Qt, y2D)

    ' 从 R 矩阵最大绝对值计算尺度感知容差
    maxAbsR = 0#
    rr0 = LBound(R, 1): rc0 = LBound(R, 2)
    For j = 1 To p
        For k = j To p
            If Abs(R(rr0 + j - 1, rc0 + k - 1)) > maxAbsR Then maxAbsR = Abs(R(rr0 + j - 1, rc0 + k - 1))
        Next k
    Next
    rTol = maxAbsR * NUM_EPS * CDbl(p)
    If rTol < RANK_TOL Then rTol = RANK_TOL  ' 绝对下限

    ' 通过回代求解 Rβ = Qᵀy
    ReDim coef(1 To p)
    For j = p To 1 Step -1
        s = Qty(LBound(Qty, 1) + j - 1, LBound(Qty, 2))
        For k = j + 1 To p
            s = s - R(rr0 + j - 1, rc0 + k - 1) * coef(k)
        Next
        If Abs(R(rr0 + j - 1, rc0 + j - 1)) < rTol Then
            coef(j) = 0#
        Else
            coef(j) = s / R(rr0 + j - 1, rc0 + j - 1)
        End If
    Next

    ' (XᵀX)⁻¹ = R⁻¹(R⁻¹)ᵀ — 从 R 直接计算，避免 XᵀX 平方条件数
    Rinv = TriangularInverse(R)
    ' 仅需 diag(R⁻¹(R⁻¹)ᵀ) 用于标准误差；显式计算 p×p 结果以复用现有 diag 提取逻辑
    RinvT = MatrixTranspose(Rinv)
    XtXinv = MatrixMultiply(Rinv, RinvT)

    ' 拟合值 ŷ = Xβ
    ReDim fitted(1 To n)
    For i = 1 To n
        fitted(i) = 0#
        For j = 1 To p
            fitted(i) = fitted(i) + X(LBound(X, 1) + i - 1, LBound(X, 2) + j - 1) * coef(j)
        Next
    Next

    ' 残差 e = y - ŷ
    ReDim resid(1 To n)
    For i = 1 To n
        resid(i) = y(LBound(y) + i - 1) - fitted(i)
    Next

    ' 总平方和, 残差平方和 — 补偿求和 (Neumaier 变体, 无分支内循环)
    ' 注: 与代码库其余部分的标准 Kahan (c/y/t) 不同, 此变体通过
    ' 有符号比较 Abs(sst) >= Abs(dy) 选择补偿项, 对平方值累加更精确
    yMean = 0#: sst = 0#: sse = 0#
    sstC = 0#: sseC = 0#
    For i = 1 To n: yMean = yMean + y(LBound(y) + i - 1): Next
    yMean = yMean / n
    For i = 1 To n
        dy = (y(LBound(y) + i - 1) - yMean) ^ 2
        dt = sst + dy
        If Abs(sst) >= Abs(dy) Then sstC = sstC + (sst - dt) + dy Else sstC = sstC + (dy - dt) + sst
        sst = dt
        dy = resid(i) ^ 2
        dt = sse + dy
        If Abs(sse) >= Abs(dy) Then sseC = sseC + (sse - dt) + dy Else sseC = sseC + (dy - dt) + sse
        sse = dt
    Next
    sst = sst + sstC
    sse = sse + sseC

    ' R², 调整 R²
    r2 = 0#: adjR2 = 0#
    If sst > 0# Then
        r2 = 1# - sse / sst
    Else
        r2 = 1#
    End If
    dfReg = p - 1
    dfRes = n - p
    If dfRes > 0 Then
        adjR2 = 1# - (1# - r2) * (n - 1) / dfRes
    ElseIf r2 >= 1# Then
        adjR2 = 1#
    End If

    ' 标准误差 SE(βⱼ) = √(σ² · diag((XᵀX)⁻¹))
    If dfRes > 0 Then sigma2 = sse / dfRes Else sigma2 = 0#
    ReDim se(1 To p)
    For j = 1 To p
        se(j) = Sqr(Abs(sigma2 * XtXinv(LBound(XtXinv, 1) + j - 1, LBound(XtXinv, 2) + j - 1)))
    Next

    ' t 统计量, p 值
    ReDim tStats(1 To p)
    ReDim pValues(1 To p)
    For j = 1 To p
        If se(j) > 0# Then
            tStats(j) = coef(j) / se(j)
        Else
            tStats(j) = 0#
        End If
        pValues(j) = ComputePValue(tStats(j), dfRes)
    Next

    ' F 检验
    fStat = 0#: fPValue = 0#
    ssReg = sst - sse
    If dfReg > 0 And dfRes > 0 Then
        If sse > 0# Then
            fStat = (ssReg / dfReg) / (sse / dfRes)
            fPValue = ComputeFPValue(fStat, dfReg, dfRes)
        Else
            fStat = 1E+300  ' 完美拟合标记
            fPValue = 0#    ' 完美拟合时 p 值为 0
        End If
    End If

    Set model = DP.Create()
    model.Add "coefficients", coef
    model.Add "fitted_values", fitted
    model.Add "residuals", resid
    model.Add "r_squared", r2
    model.Add "adj_r_squared", adjR2
    model.Add "sigma2", sigma2
    model.Add "se", se
    model.Add "t_stats", tStats
    model.Add "p_values", pValues
    model.Add "f_stat", fStat
    model.Add "f_pvalue", fPValue
    model.Add "n", n
    model.Add "p", p
    model.Add "df_residual", dfRes
    model.Add "sse", sse

    Set FitOLS = model
End Function

'=============================================================================
' ComputePValue — t 分布双尾 p 值
'=============================================================================
Private Function ComputePValue(ByVal tStat As Double, ByVal df As Long) As Double
    If df <= 0 Then ComputePValue = 1#: Exit Function
    ' 使用纯 VBA TDist2T（StatsUtils.bas），无 Excel 依赖
    ComputePValue = TDist2T(Abs(tStat), df)
End Function

Private Function ComputeFPValue(ByVal fStat As Double, ByVal df1 As Long, ByVal df2 As Long) As Double
    If df1 <= 0 Or df2 <= 0 Then ComputeFPValue = 1#: Exit Function
    ' 使用纯 VBA FDistRT（StatsUtils.bas），无 Excel 依赖
    ComputeFPValue = FDistRT(fStat, df1, df2)
End Function



'=============================================================================
' StandardizeColumns — Z-score 标准化矩阵的每一列 (每列均值0, 标准差1)
'=============================================================================
Private Sub StandardizeColumns(ByRef X() As Double, ByRef means() As Double, ByRef stds() As Double)
    Dim n As Long, p As Long, i As Long, j As Long
    Dim s As Double, ss As Double
    n = UBound(X, 1) - LBound(X, 1) + 1
    p = UBound(X, 2) - LBound(X, 2) + 1
    ReDim means(1 To p): ReDim stds(1 To p)

    For j = 1 To p
        ' Kahan 补偿求和 — 均值
        s = 0#: ss = 0#
        Dim kc As Double, ky As Double, kt As Double
        kc = 0#
        For i = 1 To n
            ky = X(i, j) - kc: kt = s + ky: kc = (kt - s) - ky: s = kt
        Next
        means(j) = s / n
        ' Kahan 补偿求和 — 方差
        ss = 0#: kc = 0#
        For i = 1 To n
            ky = (X(i, j) - means(j)) ^ 2 - kc: kt = ss + ky: kc = (kt - ss) - ky: ss = kt
        Next
        If n > 1 Then
            If ss < 0# Then ss = 0#
            stds(j) = Sqr(ss / (n - 1))
        Else
            stds(j) = 1#
        End If
        If stds(j) < RANK_TOL Then stds(j) = 1#
        For i = 1 To n: X(i, j) = (X(i, j) - means(j)) / stds(j): Next
    Next
End Sub

'=============================================================================
' EncodePredictRow — 将一行因子值编码为设计矩阵行 (用于预测)
'=============================================================================
Private Function EncodePredictRow(ByRef factorValues As Variant, _
                                   ByRef factorMap As Object, _
                                   ByVal p As Long) As Double()
    Dim row() As Double
    Dim key As Variant, colInfo As Object
    Dim fCol As Long, designCols As Variant, cf As Long
    Dim val As Variant, colType As String
    Dim levels As Variant
    Dim nLev As Long
    Dim vStr As String
    Dim catMatched As Boolean
    Dim di As Long

    ReDim row(1 To p)
    row(1) = 1#  ' 截距

    For Each key In factorMap.Keys
        Set colInfo = factorMap(key)
        fCol = CLng(key)
        colType = colInfo("type")
        designCols = colInfo("design_cols")

        If fCol >= LBound(factorValues) And fCol <= UBound(factorValues) Then
            val = factorValues(fCol)
        Else
            val = Empty
        End If

        Select Case colType
            Case "numeric"
                If UBound(designCols) >= LBound(designCols) Then
                    Err.Clear
                    On Error Resume Next
                    row(designCols(LBound(designCols))) = ToDouble(val)
                    If Err.Number <> 0 Then
                        Err.Clear
                        On Error GoTo 0
                        Err.Raise ERR_INVALID_DATA, "EncodePredictRow", _
                            "非数值 '" & CStr(val) & "' (因子列 " & fCol & ")。"
                    End If
                    On Error GoTo 0
                End If
            Case "boolean"
                If UBound(designCols) >= LBound(designCols) Then
                    Err.Clear
                    On Error Resume Next
                    row(designCols(LBound(designCols))) = ToBoolDouble(val)
                    If Err.Number <> 0 Then
                        Err.Clear
                        On Error GoTo 0
                        Err.Raise ERR_INVALID_DATA, "EncodePredictRow", _
                            "非布尔值 '" & CStr(val) & "' (因子列 " & fCol & ")。"
                    End If
                    On Error GoTo 0
                End If
            Case "categorical"
                If colInfo.Exists("levels") Then
                    levels = colInfo("levels")
                    nLev = UBound(levels) - LBound(levels) + 1
                    vStr = CStr(val)
                    If nLev > 1 Then
                        catMatched = False
                        For cf = 2 To nLev
                            di = LBound(designCols) + cf - 2
                            If vStr = CStr(levels(LBound(levels) + cf - 1)) Then
                                row(designCols(di)) = 1#
                                catMatched = True
                            Else
                                row(designCols(di)) = 0#
                            End If
                        Next
                        If Not catMatched Then
                            ' 检查是否为参考水平 (水平1在哑变量中被丢弃)
                            If vStr = CStr(levels(LBound(levels))) Then
                                ' 匹配参考水平: 所有哑变量保持为 0
                            Else
                                Err.Raise ERR_UNKNOWN_LEVEL, "EncodePredictRow", _
                                    "未知分类值 '" & vStr & "' (因子列 " & fCol & ")，不在训练数据的水平中。"
                            End If
                        End If
                    End If
                End If
        End Select
    Next

    EncodePredictRow = row
End Function

'=============================================================================
' PredictOne — 单次预测: result = x · β
'=============================================================================
Private Function PredictOne(ByRef row() As Double, ByRef coef() As Double) As Double
    Dim j As Long, s As Double
    For j = 1 To UBound(row)
        s = s + row(j) * coef(LBound(coef) + j - 1)
    Next
    PredictOne = s
End Function

'=====================================================================
' 公共函数
'=====================================================================


'=============================================================================
' FactorImportance — 因子重要性排序
'
' 基于标准化回归系数的绝对值进行排序
'
' 参数:
'   data       — Range 或 2D 数组 (首行为表头)
'   factorCols — 因子列索引数组 (1-based)
'   resultCol  — 结果列索引 (1-based)
'
' 返回: 2D Variant 数组, 列: [排名, 因子, 标准化系数, 原始系数, 绝对值]
'=============================================================================
Public Function FactorImportance(ByVal data As Variant, _
                                  ByVal factorCols As Variant, _
                                  ByVal resultCol As Long, _
                                  Optional ByVal hasHeader As Boolean = True) As Variant()
    Dim dataArr As Variant
    Dim numRows As Long, numCols As Long
    Dim firstDataRow As Long, lastDataRow As Long
    Dim n As Long
    Dim e1() As Variant, e2() As Variant
    Dim dm As Object
    Dim X() As Double, y() As Double
    Dim coefNames As Variant
    Dim factorMap As Object
    Dim p As Long
    Dim Xcopy() As Double
    Dim i As Long, j As Long, ii As Long, jj As Long, k As Long
    Dim tmp As Variant
    Dim means2() As Double, stds2() As Double
    Dim yCopy() As Double
    Dim yMean As Double, yStd As Double, yy As Double
    Dim stdCoef() As Double, rawCoef() As Double
    Dim key As Variant, colInfo As Object
    Dim fImport As Object, fRawSum As Object, fNameMap As Object
    Dim fCol As Long, dColIdx As Long, dc As Long
    Dim dColsArr As Variant
    Dim maxAbs As Double, rawBest As Double
    Dim fName As String
    Dim numFactors As Long
    Dim impData() As Variant
    Dim fIdx As Long
    Dim out() As Variant
    Dim vk As VariantKit: Set vk = New VariantKit

    dataArr = vk.NormalizeTo2D(data, numRows, numCols)

    If hasHeader Then firstDataRow = 2 Else firstDataRow = 1
    lastDataRow = numRows
    n = lastDataRow - firstDataRow + 1
    If n < 3 Then
        ReDim e1(1 To 1, 1 To 1): e1(1, 1) = "Need >=3 data rows": FactorImportance = e1: Exit Function
    End If

    Set dm = BuildDesignMatrix(dataArr, factorCols, resultCol, firstDataRow, lastDataRow, hasHeader)
    X = dm("X")
    y = dm("y")
    coefNames = dm("coef_names")
    Set factorMap = dm("factor_map")
    p = dm("p")
    n = dm("n")  ' 列表删除后的有效行数

    ' 标准化 (跳过截距列即第1列)
    ReDim Xcopy(1 To n, 1 To p)
    For i = 1 To n
        Xcopy(i, 1) = 1#
        For j = 2 To p
            Xcopy(i, j) = X(i, j)
        Next
    Next
    StandardizeColumns Xcopy, means2, stds2
    ' 恢复截距列 (未标准化 — 保持为 1)
    For i = 1 To n: Xcopy(i, 1) = 1#: Next

    ' 标准化 y
    ReDim yCopy(1 To n)
    yMean = 0#: yStd = 0#: yy = 0#
    For i = 1 To n: yMean = yMean + y(LBound(y) + i - 1): Next
    yMean = yMean / n
    For i = 1 To n: yy = y(LBound(y) + i - 1) - yMean: yStd = yStd + yy * yy: Next
    If n > 1 Then yStd = Sqr(yStd / (n - 1)) Else yStd = 1#
    If yStd < RANK_TOL Then yStd = 1#
    For i = 1 To n: yCopy(i) = (y(LBound(y) + i - 1) - yMean) / yStd: Next

    ' 拟合标准化模型 (仅需系数，使用轻量版本)
    stdCoef = FitCoefOnly(Xcopy, yCopy)

    ' 原始模型
    rawCoef = FitCoefOnly(X, y)

    ' 聚合设计列以构建每个因子的重要性。
    ' 分类因子产生 k-1 个哑变量列; 我们将其聚合
    ' 使重要性反映整个因子, 而非单个哑变量。
    Set fImport = DP.Create()
    Set fRawSum = DP.Create()
    Set fNameMap = DP.Create()

    For Each key In factorMap.Keys
        Set colInfo = factorMap(key)
        dColsArr = colInfo("design_cols")
        If IsArray(dColsArr) Then
            maxAbs = 0#
            rawBest = 0#
            For dc = LBound(dColsArr) To UBound(dColsArr)
                dColIdx = CLng(dColsArr(dc))
                If dColIdx >= 2 Then  ' skip intercept (column 1)
                    If Abs(stdCoef(LBound(stdCoef) + dColIdx - 1)) > maxAbs Then
                        maxAbs = Abs(stdCoef(LBound(stdCoef) + dColIdx - 1))
                    End If
                    ' 对于分类因子: 使用绝对值最大的原始系数
                    ' (保留符号)。将哑变量系数求和
                    ' 在统计上无意义。
                    If Abs(rawCoef(LBound(rawCoef) + dColIdx - 1)) > Abs(rawBest) Then
                        rawBest = rawCoef(LBound(rawCoef) + dColIdx - 1)
                    End If
                End If
            Next
            fName = coefNames(LBound(coefNames) + CLng(dColsArr(LBound(dColsArr))) - 1)
            If Not fImport.Exists(key) Then
                fImport.Add key, maxAbs
                fRawSum.Add key, rawBest
                fNameMap.Add key, fName
            End If
        End If
    Next

    numFactors = fImport.Count
    If numFactors < 1 Then
        ReDim e2(1 To 1, 1 To 1): e2(1, 1) = "No factors in design matrix": FactorImportance = e2: Exit Function
    End If

    ' 构建未排序表格 (每个因子一行, 而非每个哑变量列一行)
    ReDim impData(1 To numFactors, 1 To 5)
    fIdx = 0
    For Each key In fImport.Keys
        fIdx = fIdx + 1
        impData(fIdx, 1) = 0#
        impData(fIdx, 2) = fNameMap(key)       ' factor name
        impData(fIdx, 3) = CDbl(fImport(key))   ' max abs stdCoef — aggregated
        impData(fIdx, 4) = CDbl(fRawSum(key))   ' sum of raw coefs
        impData(fIdx, 5) = CDbl(fImport(key))   ' sort key
    Next

    ' QuickSortIndicesByScore 已用于 OptimizeFactors，此处复用同一下行路径
    Dim scores() As Double, idx() As Long
    ReDim scores(1 To numFactors)
    ReDim idx(1 To numFactors)
    For ii = 1 To numFactors
        scores(ii) = CDbl(impData(ii, 5))
        idx(ii) = ii
    Next
    QuickSortIndicesByScore scores, idx, 1, numFactors
    ' 按排序后的索引重排 impData
    Dim sorted() As Variant
    ReDim sorted(1 To numFactors, 1 To 5)
    For ii = 1 To numFactors
        For k = 1 To 5
            sorted(ii, k) = impData(idx(ii), k)
        Next
        sorted(ii, 1) = ii  ' 排名
    Next
    impData = sorted

    ' 构建返回数组 (含表头)
    ReDim out(1 To numFactors + 1, 1 To 5)
    out(1, 1) = "排名": out(1, 2) = "因子": out(1, 3) = "标准化系数"
    out(1, 4) = "原始系数": out(1, 5) = "绝对重要性"
    For ii = 1 To numFactors
        For k = 1 To 5: out(ii + 1, k) = impData(ii, k): Next
    Next

    FactorImportance = out
End Function

'=============================================================================
' InteractionEffects — 检测因子间的交互效应
'
' 对每一对因子，在模型中加入交互项，检验交互项的联合显著性 (F检验)。
' 支持数值/布尔/分类因子的任意组合。
'
' 参数:
'   data       — Range 或 2D 数组 (首行为表头)
'   factorCols — 因子列索引数组 (1-based)
'   resultCol  — 结果列索引
'
' 返回: 2D Variant 数组, 列: [因子A, 因子B, 交互项数, F值, p值, 是否显著(p<0.05)]
'=============================================================================
Public Function InteractionEffects(ByVal data As Variant, _
                                    ByVal factorCols As Variant, _
                                    ByVal resultCol As Long, _
                                    Optional ByVal hasHeader As Boolean = True) As Variant()
    Const MAX_INTERACT_TERMS As Long = 50

    Dim dataArr As Variant
    Dim numRows As Long, numCols As Long
    Dim firstDataRow As Long, lastDataRow As Long
    Dim n As Long
    Dim e1() As Variant, e2() As Variant
    Dim dm As Object
    Dim X() As Double, y() As Double
    Dim coefNames As Variant
    Dim p As Long
    Dim factorMap As Object
    Dim key As Variant, colInfo As Object, dCols As Variant
    Dim fi As Long, fj As Long
    Dim factorKeys As Variant, nf As Long
    Dim numPairs As Long
    Dim intResults() As Variant
    Dim pairIdx As Long
    Dim nColsA As Long, nColsB As Long
    Dim colsA() As Long, colsB() As Long
    Dim i As Long
    Dim Xint() As Double, pInt As Long, intCount As Long
    Dim ci As Long, cj As Long, colIdx As Long
    Dim intModel As Object, baseModel As Object
    Dim sseBase As Double, sseFull As Double
    Dim dfDiff As Long, dfFull As Long
    Dim fStat As Double, pVal As Double
    Dim vk As VariantKit: Set vk = New VariantKit

    dataArr = vk.NormalizeTo2D(data, numRows, numCols)

    If hasHeader Then firstDataRow = 2 Else firstDataRow = 1
    lastDataRow = numRows
    n = lastDataRow - firstDataRow + 1
    If n < 5 Then
        ReDim e1(1 To 1, 1 To 1): e1(1, 1) = "Need >=5 data rows": InteractionEffects = e1: Exit Function
    End If

    Set dm = BuildDesignMatrix(dataArr, factorCols, resultCol, firstDataRow, lastDataRow, hasHeader)
    X = dm("X")
    y = dm("y")
    coefNames = dm("coef_names")
    p = dm("p")
    Set factorMap = dm("factor_map")
    n = dm("n")

    ' 基准模型 (不含交互项)
    Set baseModel = FitOLS(X, y)
    sseBase = CDbl(baseModel("sse"))

    ' 收集所有因子 (数值/布尔/分类均参与交互检测)
    factorKeys = factorMap.Keys
    nf = factorMap.Count
    Dim fkLb As Long: fkLb = LBound(factorKeys)

    If nf < 2 Then
        ReDim e2(1 To 1, 1 To 1): e2(1, 1) = "Need >=2 factors for interactions": InteractionEffects = e2: Exit Function
    End If

    numPairs = nf * (nf - 1) \ 2
    ReDim intResults(1 To numPairs + 1, 1 To 6)
    intResults(1, 1) = "Factor A": intResults(1, 2) = "Factor B"
    intResults(1, 3) = "Interaction Terms": intResults(1, 4) = "F Value": intResults(1, 5) = "p Value": intResults(1, 6) = "Significant"

    pairIdx = 0
    For fi = 1 To nf - 1
        Set colInfo = factorMap(CStr(factorKeys(fkLb + fi - 1)))
        dCols = colInfo("design_cols")
        If Not IsArray(dCols) Then GoTo NextFi
        nColsA = UBound(dCols) - LBound(dCols) + 1
        ReDim colsA(1 To nColsA)
        For ci = 1 To nColsA
            colsA(ci) = CLng(dCols(LBound(dCols) + ci - 1))
        Next

        For fj = fi + 1 To nf
            Set colInfo = factorMap(CStr(factorKeys(fkLb + fj - 1)))
            dCols = colInfo("design_cols")
            If Not IsArray(dCols) Then GoTo NextFj
            nColsB = UBound(dCols) - LBound(dCols) + 1
            intCount = nColsA * nColsB
            If intCount > MAX_INTERACT_TERMS Then
                pairIdx = pairIdx + 1
                intResults(pairIdx + 1, 1) = coefNames(LBound(coefNames) + colsA(1) - 1)
                intResults(pairIdx + 1, 2) = coefNames(LBound(coefNames) + dCols(LBound(dCols)) - 1)
                intResults(pairIdx + 1, 3) = ">" & MAX_INTERACT_TERMS
                intResults(pairIdx + 1, 4) = Empty: intResults(pairIdx + 1, 5) = Empty
                intResults(pairIdx + 1, 6) = "Skipped"
                GoTo NextFj
            End If

            ReDim colsB(1 To nColsB)
            For cj = 1 To nColsB
                colsB(cj) = CLng(dCols(LBound(dCols) + cj - 1))
            Next

            pairIdx = pairIdx + 1

            ' 构建扩展设计矩阵: X + 所有交互列
            pInt = p + intCount
            ReDim Xint(1 To n, 1 To pInt)
            For i = 1 To n
                For cj = 1 To p
                    Xint(i, cj) = X(i, cj)
                Next
                colIdx = p
                For ci = 1 To nColsA
                    For cj = 1 To nColsB
                        colIdx = colIdx + 1
                        Xint(i, colIdx) = X(i, colsA(ci)) * X(i, colsB(cj))
                    Next
                Next
            Next

            ' 拟合全模型并 F 检验
            On Error Resume Next
            Set intModel = FitOLS(Xint, y)
            If Err.Number <> 0 Then
                Err.Clear: On Error GoTo 0
                intResults(pairIdx + 1, 1) = coefNames(LBound(coefNames) + colsA(1) - 1)
                intResults(pairIdx + 1, 2) = coefNames(LBound(coefNames) + colsB(1) - 1)
                intResults(pairIdx + 1, 3) = intCount
                intResults(pairIdx + 1, 4) = "Singular": intResults(pairIdx + 1, 5) = "Singular"
                intResults(pairIdx + 1, 6) = "—"
                GoTo NextFj
            End If
            On Error GoTo 0

            sseFull = CDbl(intModel("sse"))
            dfFull = CLng(intModel("df_residual"))
            dfDiff = pInt - p

            If dfFull > 0 And sseFull > 0# Then
                fStat = ((sseBase - sseFull) / dfDiff) / (sseFull / dfFull)
                If fStat < 0# Then fStat = 0#
            Else
                fStat = 0#
            End If
            pVal = ComputeFPValue(fStat, dfDiff, dfFull)

            intResults(pairIdx + 1, 1) = coefNames(LBound(coefNames) + colsA(1) - 1)
            intResults(pairIdx + 1, 2) = coefNames(LBound(coefNames) + colsB(1) - 1)
            intResults(pairIdx + 1, 3) = intCount
            intResults(pairIdx + 1, 4) = fStat
            intResults(pairIdx + 1, 5) = pVal
            If pVal < 0.05 Then intResults(pairIdx + 1, 6) = "Yes" Else intResults(pairIdx + 1, 6) = "No"

NextFj:
        Next fj
NextFi:
    Next fi

    ' 截断至实际对数
    If pairIdx < numPairs Then
        ReDim Preserve intResults(1 To pairIdx + 1, 1 To 6)
    End If

    InteractionEffects = intResults
End Function

'=============================================================================
' ANOVAOneWay — 单因素方差分析
'
' 检验一个分类因子对一个数值结果是否有显著影响
'
' 参数:
'   data      — Range 或 2D 数组 (首行为表头)
'   factorCol — 分类因子列索引 (1-based)
'   resultCol — 数值结果列索引 (1-based)
'
' 返回: Dictionary {"SSB", "SSW", "SST", "dfB", "dfW", "dfT",
'                    "MSB", "MSW", "F", "p_value", "eta_sq", "summary"}
'=============================================================================
Public Function ANOVAOneWay(ByVal data As Variant, _
                             ByVal factorCol As Long, _
                             ByVal resultCol As Long, _
                             Optional ByVal hasHeader As Boolean = True) As Object
    Dim dataArr As Variant
    Dim numRows As Long, numCols As Long
    Dim firstDataRow As Long
    Dim n As Long
    Dim groups As Object
    Dim r As Long, v As Variant, key As Variant
    Dim allValues() As Double
    Dim grandSum As Double
    Dim idx As Long
    Dim grp As Object
    Dim val As Double
    Dim result As Object
    Dim k As Long
    Dim grandMean As Double
    Dim ssb As Double, ssw As Double
    Dim groupMean As Double
    Dim sst As Double
    Dim dfB As Long, dfW As Long, dfT As Long
    Dim msb As Double, msw As Double, fStat As Double, pVal As Double
    Dim etaSq As Double
    Dim rDict As Object
    Dim summary As String
    Dim vk As VariantKit: Set vk = New VariantKit

    dataArr = vk.NormalizeTo2D(data, numRows, numCols)
    ' NormalizeTo2D may return 0-based COM arrays — convert to 1-based
    Dim lbR As Long: lbR = LBound(dataArr, 1)
    Dim lbC As Long: lbC = LBound(dataArr, 2)
    If lbR <> 1 Or lbC <> 1 Then
        Dim normArr() As Variant, ri As Long, ci As Long
        numRows = UBound(dataArr, 1) - lbR + 1
        numCols = UBound(dataArr, 2) - lbC + 1
        ReDim normArr(1 To numRows, 1 To numCols)
        For ri = 1 To numRows
            For ci = 1 To numCols
                normArr(ri, ci) = dataArr(lbR + ri - 1, lbC + ci - 1)
            Next ci
        Next ri
        dataArr = normArr
    End If
    If hasHeader Then firstDataRow = 2 Else firstDataRow = 1
    n = numRows - firstDataRow + 1

    ' 收集各水平的数据
    Set groups = DP.Create()
    For r = firstDataRow To numRows
        key = CStr(dataArr(r, factorCol))
        If Not groups.Exists(key) Then
            groups.Add key, DP.Create()
            groups(key).Add "count", 0
            groups(key).Add "sum", 0#
            groups(key).Add "sumsq", 0#
        End If
    Next

    ' 计算组内统计
    ReDim allValues(1 To n)
    grandSum = 0#
    idx = 0

    For r = firstDataRow To numRows
        If Not IsNumericCell(dataArr(r, resultCol)) Then
            ' 跳过结果列缺失的行
        Else
            key = CStr(dataArr(r, factorCol))
            val = CDbl(dataArr(r, resultCol))
            idx = idx + 1
            allValues(idx) = val
            grandSum = grandSum + val
            Set grp = groups(key)
            grp("count") = grp("count") + 1
            grp("sum") = grp("sum") + val
            grp("sumsq") = grp("sumsq") + val * val
        End If
    Next
    n = idx  ' 更新为有效行数 (列表删除)

    If n = 0 Then
        Set result = DP.Create()
        result.Add "error", "No valid numeric data in result column"
        Set ANOVAOneWay = result: Exit Function
    End If

    ' 统计有效组数 (排除 count=0 的空组)
    k = 0
    For Each key In groups.Keys
        If groups(key)("count") > 0 Then k = k + 1
    Next
    If k < 2 Then
        Set result = DP.Create()
        result.Add "error", "Need at least 2 groups with valid data"
        Set ANOVAOneWay = result: Exit Function
    End If

    grandMean = grandSum / n

    ' 组间平方和 — 存储组均值用于两阶段组内平方和计算
    ssb = 0#: ssw = 0#
    For Each key In groups.Keys
        Set grp = groups(key)
        If grp("count") > 0 Then
            groupMean = grp("sum") / grp("count")
        Else
            groupMean = 0#
        End If
        ssb = ssb + grp("count") * (groupMean - grandMean) ^ 2
        grp("mean") = groupMean
    Next

    ' 两阶段组内平方和: 避免 sumsq - n*mean² 导致的灾难性抵消
    For r = firstDataRow To numRows
        If Not IsNumericCell(dataArr(r, resultCol)) Then
            ' 跳过结果列缺失的行
        Else
            key = CStr(dataArr(r, factorCol))
            val = CDbl(dataArr(r, resultCol))
            Set grp = groups(key)
            ssw = ssw + (val - grp("mean")) * (val - grp("mean"))
        End If
    Next

    sst = ssb + ssw

    dfB = k - 1
    dfW = n - k
    dfT = n - 1

    msb = 0#: msw = 0#: fStat = 0#: pVal = 0#
    If dfW > 0 Then
        msb = ssb / dfB
        msw = ssw / dfW
        If msw > 0# Then fStat = msb / msw
        pVal = ComputeFPValue(fStat, dfB, dfW)
    End If

    If sst > 0# Then etaSq = ssb / sst

    Set rDict = DP.Create()
    rDict.Add "SSB", ssb: rDict.Add "SSW", ssw: rDict.Add "SST", sst
    rDict.Add "dfB", dfB: rDict.Add "dfW", dfW: rDict.Add "dfT", dfT
    rDict.Add "MSB", msb: rDict.Add "MSW", msw
    rDict.Add "F", fStat: rDict.Add "p_value", pVal
    rDict.Add "eta_sq", etaSq
    rDict.Add "n_groups", k: rDict.Add "n_total", n
    rDict.Add "significant", (pVal < 0.05)

    Dim pStr As String: If pVal < 0.0001 Then pStr = Format(pVal, "0.0000E+00") Else pStr = Format(pVal, "0.0000")
    summary = "F(" & dfB & "," & dfW & ") = " & Format(fStat, "0.000000") & _
              ", p = " & pStr & _
              ", eta² = " & Format(etaSq, "0.000")
    rDict.Add "summary", summary

    Set ANOVAOneWay = rDict
End Function

' Wrapper for COM testing — returns just the F-statistic as Double
' Returns -1# on error (insufficient groups, invalid input, etc.)
Public Function ANOVAOneWay_Fstat(ByVal data As Variant, _
                                   ByVal factorCol As Long, _
                                   ByVal resultCol As Long, _
                                   Optional ByVal hasHeader As Boolean = True) As Double
    Dim result As Object
    Set result = ANOVAOneWay(data, factorCol, resultCol, hasHeader)
    If result Is Nothing Then ANOVAOneWay_Fstat = -1#: Exit Function
    If result.Exists("error") Then ANOVAOneWay_Fstat = -1#: Exit Function
    If result.Exists("F") Then ANOVAOneWay_Fstat = CDbl(result("F")) Else ANOVAOneWay_Fstat = -1#
End Function

'=============================================================================
' LinearModelFit — 多元线性回归拟合
'
' 参数:
'   data       — Range 或 2D 数组 (首行为表头)
'   factorCols — 因子列索引数组 (1-based)
'   resultCol  — 结果列索引 (1-based)
'
' 返回: Dictionary (模型对象)
'   "coefficients", "coef_names", "r_squared", "adj_r_squared",
'   "se", "t_stats", "p_values", "f_stat", "f_pvalue",
'   "fitted_values", "residuals", "n", "p", "df_residual",
'   "factor_map" (用于预测), "formula" (公式文本)
'=============================================================================
Public Function LinearModelFit(ByVal data As Variant, _
                                ByVal factorCols As Variant, _
                                ByVal resultCol As Long, _
                                Optional ByVal hasHeader As Boolean = True) As Object
    Dim dataArr As Variant
    Dim numRows As Long, numCols As Long
    Dim firstDataRow As Long
    Dim n As Long
    Dim dm As Object
    Dim X() As Double, y() As Double
    Dim model As Object
    Dim formula As String
    Dim fn As Variant
    Dim fi As Long
    Dim parts() As String
    Dim vk As VariantKit: Set vk = New VariantKit

    dataArr = vk.NormalizeTo2D(data, numRows, numCols)
    If hasHeader Then firstDataRow = 2 Else firstDataRow = 1
    n = numRows - firstDataRow + 1
    If n < 3 Then
        Err.Raise ERR_TOO_FEW_ROWS, "LinearModelFit", _
            "需要至少 3 行数据（含表头），当前仅有 " & n & " 行。"
    End If

    Set dm = BuildDesignMatrix(dataArr, factorCols, resultCol, firstDataRow, numRows, hasHeader)
    X = dm("X")
    y = dm("y")

    Set model = FitOLS(X, y)
    model.Add "coef_names", dm("coef_names")
    model.Add "factor_map", dm("factor_map")

    ' 构建公式文本
    If hasHeader Then
        formula = CStr(dataArr(1, resultCol)) & " ~ "
    Else
        formula = "Y ~ "
    End If
    fn = dm("coef_names")
    ReDim parts(1 To UBound(fn) - 1)
    For fi = 2 To UBound(fn)
        parts(fi - 1) = fn(LBound(fn) + fi - 1)
    Next
    formula = formula & Join(parts, " + ")
    model.Add "formula", formula
    ' 存储原始因子列顺序用于预测
    model.Add "factor_cols", factorCols

    Set LinearModelFit = model
End Function

'=============================================================================
' LinearModelPredict — 基于模型预测结果
'
' 参数:
'   model   — 由 LinearModelFit 返回的模型 Dictionary
'   newData — 新因子值: 1D 数组 (按 factorCols 顺序) 或单行 Range
'
' 返回: 预测值 (Double)
'=============================================================================
Public Function LinearModelPredict(ByVal model As Object, ByVal newData As Variant) As Double
    Dim factorMap As Object: Set factorMap = model("factor_map")
    Dim coef() As Double: coef = model("coefficients")
    Dim p As Long: p = CLng(model("p"))
    Dim factorCols As Variant
    Dim nf As Long
    Dim factorValues() As Variant
    Dim ci As Long, maxCol As Long, fi As Long
    Dim rng As Range
    Dim newDataIdx As Long
    Dim row() As Double

    ' 使用存储的 factorCols 顺序进行稳定的预测映射
    If model.Exists("factor_cols") Then
        factorCols = model("factor_cols")
    Else
        factorCols = factorMap.Keys  ' 向后兼容
    End If
    nf = UBound(factorCols) - LBound(factorCols) + 1

    maxCol = 0
    For fi = 0 To nf - 1
        ci = CLng(factorCols(LBound(factorCols) + fi))
        If ci > maxCol Then maxCol = ci
    Next
    ReDim factorValues(1 To maxCol)

    If IsObject(newData) Then
        If TypeOf newData Is Range Then
            Set rng = newData
            If rng.Columns.Count > 1 Then Set rng = rng.Rows(1)
            For fi = 0 To nf - 1
                ci = CLng(factorCols(LBound(factorCols) + fi))
                If fi + 1 <= rng.Cells.Count Then
                    factorValues(ci) = rng.Cells(1, fi + 1).Value
                End If
            Next
        End If
    ElseIf IsArray(newData) Then
        For fi = 0 To nf - 1
            ci = CLng(factorCols(LBound(factorCols) + fi))
            newDataIdx = LBound(newData) + fi
            If newDataIdx <= UBound(newData) Then
                factorValues(ci) = newData(newDataIdx)
            End If
        Next
    End If

    row = EncodePredictRow(factorValues, factorMap, p)
    LinearModelPredict = PredictOne(row, coef)
End Function

'=============================================================================
' FactorSweep — 单因子 What-If 扫描
'
' 保持其他因子不变，扫描一个因子的变化对结果的影响
'
' 参数:
'   model       — 由 LinearModelFit 返回的模型 Dictionary
'   factorIndex — 要扫描的因子序号 (在 factorCols 中的 1-based 序号)
'   fromVal     — 扫描起始值
'   toVal       — 扫描结束值
'   steps       — 扫描步数
'   baseValues  — 其他因子的基准值 (1D 数组, 按 factorCols 顺序;
'                  省略时数值/布尔因子取0, 分类因子取第一个水平)
'
' 返回: 2D Variant 数组, 列: [因子值, 预测结果]
'=============================================================================
Public Function FactorSweep(ByVal model As Object, _
                             ByVal factorIndex As Long, _
                             ByVal fromVal As Double, _
                             ByVal toVal As Double, _
                             ByVal steps As Long, _
                             Optional ByVal baseValues As Variant) As Variant()
    Dim factorMap As Object
    Dim coef() As Double
    Dim p As Long
    Dim nf As Long
    Dim maxCol As Long
    Dim key As Variant, ci As Long
    Dim baseArr() As Variant
    Dim bvIdx As Long
    Dim fKeys As Variant
    Dim colInfo As Object
    Dim origCol As Long
    Dim levs As Variant
    Dim e2() As Variant
    Dim out() As Variant
    Dim si As Long
    Dim targetCol As Long
    Dim val As Double
    Dim row() As Double

    Set factorMap = model("factor_map")
    coef = model("coefficients")
    p = CLng(model("p"))
    nf = factorMap.Count

    maxCol = 0
    For Each key In factorMap.Keys
        ci = CLng(key)
        If ci > maxCol Then maxCol = ci
    Next

    ' 构建基准值 (baseValues 按 factorCols 顺序, 映射到原始列索引)
    ReDim baseArr(LBound(coef) To maxCol)
    If Not IsMissing(baseValues) And IsArray(baseValues) Then
        bvIdx = LBound(baseValues)
        For Each key In factorMap.Keys
            ci = CLng(key)
            If bvIdx <= UBound(baseValues) Then
                baseArr(ci) = baseValues(bvIdx)
                bvIdx = bvIdx + 1
            End If
        Next
    End If

    fKeys = factorMap.Keys
    Dim fkLb As Long: fkLb = LBound(fKeys)

    ' 为缺失的基准值填充默认值
    For ci = fkLb To fkLb + nf - 1
        origCol = CLng(fKeys(ci))
        If IsEmpty(baseArr(origCol)) Then
            Set colInfo = factorMap(CStr(origCol))
            Select Case colInfo("type")
                Case "numeric", "boolean"
                    If colInfo.Exists("mean") Then baseArr(origCol) = colInfo("mean") Else baseArr(origCol) = 0#
                Case "categorical"
                    If colInfo.Exists("levels") Then
                        levs = colInfo("levels")
                        baseArr(origCol) = levs(LBound(levs))
                    End If
            End Select
        End If
    Next

    ' 生成扫描
    If factorIndex < 1 Or factorIndex > nf Then
        ReDim e2(1 To 1, 1 To 1): e2(1, 1) = "factorIndex out of range": FactorSweep = e2: Exit Function
    End If
    If steps < 0 Then
        ReDim e2(1 To 1, 1 To 1): e2(1, 1) = "steps must be >= 0": FactorSweep = e2: Exit Function
    End If

    ReDim out(1 To steps + 1, 1 To 2)
    out(1, 1) = "FactorValue": out(1, 2) = "Predicted"

    targetCol = CLng(fKeys(fkLb + factorIndex - 1))
    For si = 1 To steps
        If steps > 1 Then
            val = fromVal + (toVal - fromVal) * (si - 1) / (steps - 1)
        Else
            val = fromVal
        End If
        baseArr(targetCol) = val
        row = EncodePredictRow(baseArr, factorMap, p)
        out(si + 1, 1) = val
        out(si + 1, 2) = PredictOne(row, coef)
    Next

    FactorSweep = out
End Function

'=============================================================================
' OptimizeFactors — 网格搜索最优因子组合
'
' 参数:
'   data       — Range 或 2D 数组 (首行为表头)
'   factorCols — 因子列索引数组 (1-based)
'   resultCol  — 结果列索引
'   goal       — "max" 最大化, "min" 最小化, 或数值 (接近目标值)
'   topN       — 返回最优组合数 (默认 10)
'   nSteps     — 数值因子的离散步数 (默认 10)
'
' 返回: 2D Variant 数组, 每行为一个最优因子组合 + 预测结果
'=============================================================================
Public Function OptimizeFactors(ByVal data As Variant, _
                                 ByVal factorCols As Variant, _
                                 ByVal resultCol As Long, _
                                 Optional ByVal goal As Variant = "max", _
                                 Optional ByVal topN As Long = 10, _
                                 Optional ByVal nSteps As Long = 10, _
                                 Optional ByVal hasHeader As Boolean = True) As Variant()
    Dim dataArr As Variant
    Dim numRows As Long, numCols As Long
    Dim firstDataRow As Long, lastDataRow As Long
    Dim model As Object
    Dim factorMap As Object
    Dim coef() As Double
    Dim p As Long
    Dim fKeys As Variant
    Dim nf As Long
    Dim searchSpace() As Variant
    Dim totalCombos As Double
    Dim fi As Long, origCol As Long
    Dim colInfo As Object
    Dim numVals() As Double
    Dim vMin As Double, vMax As Double
    Dim r As Long, nv As Double
    Dim si As Long
    Dim boolVals() As Double
    Dim catVals() As String
    Dim e2() As Variant
    Dim resultCount As Long
    Dim bestVals() As Double
    Dim bestRows() As Variant
    Dim goalMode As String
    Dim indices() As Long
    Dim maxCol As Long
    Dim currentVals() As Variant
    Dim done As Boolean
    Dim row2() As Double
    Dim pred As Double
    Dim score As Double
    Dim worstIdx As Long, worstVal As Double
    Dim ti As Long
    Dim carry As Long
    Dim sz As Long
    Dim sortIdx() As Long
    Dim ii As Long, kk As Long
    Dim tmpScores() As Double
    Dim tmpRows() As Variant
    Dim out() As Variant

    If nSteps < 2 Then nSteps = 2
    If nSteps > 50 Then nSteps = 50
    If topN < 1 Then topN = 1
    If topN > 100 Then topN = 100
    Dim vk As VariantKit: Set vk = New VariantKit

    dataArr = vk.NormalizeTo2D(data, numRows, numCols)
    If hasHeader Then firstDataRow = 2 Else firstDataRow = 1
    lastDataRow = numRows

    ' 拟合模型
    Set model = LinearModelFit(data, factorCols, resultCol, hasHeader)
    Set factorMap = model("factor_map")
    coef = model("coefficients")
    p = CLng(model("p"))

    ' 解析每个因子的搜索空间
    fKeys = factorMap.Keys
    nf = factorMap.Count
    Dim ofkLb As Long: ofkLb = LBound(fKeys)

    ' 存储搜索空间: 每个因子一个数组
    ReDim searchSpace(1 To nf)
    ' Double 可精确表示最大 2^53≈9e15 的整数, 对 MAX_GRID_COMBOS=200k 安全
    totalCombos = 1#

    For fi = 1 To nf
        origCol = CLng(fKeys(ofkLb + fi - 1))
        Set colInfo = factorMap(CStr(origCol))

        Select Case colInfo("type")
            Case "numeric"
                vMin = 1E+300: vMax = -1E+300
                For r = firstDataRow To lastDataRow
                    If IsNumericCell(dataArr(r, origCol)) Then
                        nv = CDbl(dataArr(r, origCol))
                        If nv < vMin Then vMin = nv
                        If nv > vMax Then vMax = nv
                    End If
                Next
                ' 若无有效数据则使用默认范围
                If vMin > vMax Then
                    vMin = 0#: vMax = 1#
                End If
                If vMax - vMin < RANK_TOL Then vMax = vMin + 1#
                ReDim numVals(1 To nSteps)
                For si = 1 To nSteps
                    numVals(si) = vMin + (vMax - vMin) * (si - 1) / (nSteps - 1)
                Next
                searchSpace(fi) = numVals
                totalCombos = totalCombos * nSteps

            Case "boolean"
                ReDim boolVals(1 To 2)
                boolVals(1) = 0#: boolVals(2) = 1#
                searchSpace(fi) = boolVals
                totalCombos = totalCombos * 2#

            Case "categorical"
                If colInfo.Exists("levels") Then
                    catVals = colInfo("levels")
                Else
                    ReDim catVals(0 To 0): catVals(0) = ""
                End If
                searchSpace(fi) = catVals
                totalCombos = totalCombos * (UBound(catVals) - LBound(catVals) + 1)
        End Select
    Next

    If totalCombos > MAX_GRID_COMBOS Then
        ReDim e2(1 To 1, 1 To 1)
        e2(1, 1) = "Grid too large: " & Format(totalCombos, "#,##0") & " combos > " & MAX_GRID_COMBOS & ". Reduce nSteps."
        OptimizeFactors = e2: Exit Function
    End If

    ' 网格搜索
    resultCount = 0
    ReDim bestVals(1 To topN)
    ReDim bestRows(1 To topN, 1 To nf + 1)
    If IsNumeric(goal) Then
        goalMode = "target"
    Else
        goalMode = LCase(CStr(goal))
    End If

    ' 当前组合
    ReDim indices(1 To nf)
    For fi = 1 To nf: indices(fi) = 1: Next

    maxCol = 0
    For fi = 1 To nf
        If CLng(fKeys(ofkLb + fi - 1)) > maxCol Then maxCol = CLng(fKeys(ofkLb + fi - 1))
    Next

    ReDim currentVals(1 To maxCol)

    done = False
    Do While Not done
        ' 构建当前因子值
        For fi = 1 To nf
            origCol = CLng(fKeys(ofkLb + fi - 1))
            Set colInfo = factorMap(CStr(origCol))
            Select Case colInfo("type")
                Case "numeric", "boolean"
                    currentVals(origCol) = searchSpace(fi)(indices(fi))
                Case "categorical"
                    currentVals(origCol) = searchSpace(fi)(LBound(searchSpace(fi)) + indices(fi) - 1)
            End Select
        Next

        ' 预测
        row2 = EncodePredictRow(currentVals, factorMap, p)
        pred = PredictOne(row2, coef)

        ' 评估
        Select Case goalMode
            Case "max": score = pred
            Case "min": score = -pred
            Case "target": score = -Abs(pred - CDbl(goal))
        End Select

        ' 维护 topN
        If resultCount < topN Then
            resultCount = resultCount + 1
            bestVals(resultCount) = score
            For fi = 1 To nf: bestRows(resultCount, fi) = currentVals(CLng(fKeys(ofkLb + fi - 1))): Next
            bestRows(resultCount, nf + 1) = pred
        Else
            ' 找最差值替换
            worstIdx = 1
            worstVal = bestVals(1)
            For ti = 2 To topN
                If bestVals(ti) < worstVal Then
                    worstVal = bestVals(ti)
                    worstIdx = ti
                End If
            Next
            If score > worstVal Then
                bestVals(worstIdx) = score
                For fi = 1 To nf: bestRows(worstIdx, fi) = currentVals(CLng(fKeys(ofkLb + fi - 1))): Next
                bestRows(worstIdx, nf + 1) = pred
            End If
        End If

        ' 递增索引
        carry = 1
        For fi = nf To 1 Step -1
            indices(fi) = indices(fi) + 1
            If IsArray(searchSpace(fi)) Then
                sz = UBound(searchSpace(fi)) - LBound(searchSpace(fi)) + 1
            Else
                sz = 0
            End If
            If indices(fi) <= sz Then
                carry = 0
                Exit For
            End If
            indices(fi) = 1
        Next
        If carry = 1 Then done = True
    Loop

    ' 排序 bestRows (通过索引数组的 QuickSort, 按分数降序)
    If resultCount > 1 Then
        ReDim sortIdx(1 To resultCount)
        For ii = 1 To resultCount: sortIdx(ii) = ii: Next
        QuickSortIndicesByScore bestVals, sortIdx, 1, resultCount
        ' 使用排序后的索引重排 (通过临时副本原地重排)
        tmpScores = bestVals
        tmpRows = bestRows
        For ii = 1 To resultCount
            bestVals(ii) = tmpScores(sortIdx(ii))
            For kk = 1 To nf + 1
                bestRows(ii, kk) = tmpRows(sortIdx(ii), kk)
            Next
        Next
    End If

    ' 构建输出 (含表头)
    ReDim out(1 To resultCount + 1, 1 To nf + 1)
    For fi = 1 To nf
        If hasHeader Then
            out(1, fi) = CStr(dataArr(1, CLng(fKeys(ofkLb + fi - 1))))
        Else
            out(1, fi) = "F" & fi
        End If
    Next
    out(1, nf + 1) = "Predicted"

    For ii = 1 To resultCount
        For fi = 1 To nf + 1
            out(ii + 1, fi) = bestRows(ii, fi)
        Next
    Next

    OptimizeFactors = out
End Function

'=====================================================================
' 内部辅助
'=====================================================================
' QuickSortIndicesByScore — Hybrid QuickSort+InsertionSort for index arrays (in-place, descending by score)
' Sorts idx(lo..hi) by vals(idx(i)) descending. Recurses on smaller partition first → O(log n) stack.
'=====================================================================

Private Sub QuickSortIndicesByScore(ByRef vals() As Double, ByRef idx() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, lo2 As Long, hi2 As Long, pivot As Double, tmp As Long
    Dim key As Long

    If hi - lo <= REG_IDX_THRESHOLD Then
        For i = lo + 1 To hi
            key = idx(i): j = i - 1
            Do While j >= lo
                If vals(idx(j)) >= vals(key) Then Exit Do
                idx(j + 1) = idx(j): j = j - 1
            Loop
            idx(j + 1) = key
        Next i
        Exit Sub
    End If

    lo2 = lo: hi2 = hi
    pivot = vals(idx((lo + hi) \ 2))
    Do While lo2 <= hi2
        Do While vals(idx(lo2)) > pivot: lo2 = lo2 + 1: Loop
        Do While vals(idx(hi2)) < pivot: hi2 = hi2 - 1: Loop
        If lo2 <= hi2 Then
            tmp = idx(lo2): idx(lo2) = idx(hi2): idx(hi2) = tmp
            lo2 = lo2 + 1: hi2 = hi2 - 1
        End If
    Loop

    If hi2 - lo < hi - lo2 Then
        If lo < hi2 Then QuickSortIndicesByScore vals, idx, lo, hi2
        If lo2 < hi Then QuickSortIndicesByScore vals, idx, lo2, hi
    Else
        If lo2 < hi Then QuickSortIndicesByScore vals, idx, lo2, hi
        If lo < hi2 Then QuickSortIndicesByScore vals, idx, lo, hi2
    End If
End Sub



'=====================================================================
' UDF 包装函数 (Excel 工作表调用)
'=====================================================================

Private Function ParseFactorCols(ByVal factorCols As Range) As Variant()
    Dim fc() As Variant
    Dim cell As Range, i As Long
    ReDim fc(1 To factorCols.Cells.Count)
    i = 1
    For Each cell In factorCols.Cells
        fc(i) = CLng(cell.Value)
        i = i + 1
    Next
    ParseFactorCols = fc
End Function

Public Function UDF_REGRESS_CORREL(ByVal data As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_REGRESS_CORREL = CorrelationMatrix(data)  ' implementation in StatsUtils.bas
    Exit Function
ErrHandler:
    UDF_REGRESS_CORREL = CVErr(xlErrValue)
End Function

Public Function UDF_REGRESS_IMPORTANCE(ByVal data As Variant, ByVal factorCols As Variant, ByVal resultCol As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_REGRESS_IMPORTANCE = FactorImportance(data, ParseFactorCols(factorCols), resultCol)
    Exit Function
ErrHandler:
    UDF_REGRESS_IMPORTANCE = CVErr(xlErrValue)
End Function

Public Function UDF_REGRESS_INTERACT(ByVal data As Variant, ByVal factorCols As Variant, ByVal resultCol As Variant) As Variant
    On Error GoTo ErrHandler
    UDF_REGRESS_INTERACT = InteractionEffects(data, ParseFactorCols(factorCols), resultCol)
    Exit Function
ErrHandler:
    UDF_REGRESS_INTERACT = CVErr(xlErrValue)
End Function

Public Function UDF_REGRESS_ANOVA(ByVal data As Variant, ByVal factorCol As Variant, ByVal resultCol As Variant) As Variant
    On Error GoTo ErrHandler
    Dim anovaResult As Object: Set anovaResult = ANOVAOneWay(data, factorCol, resultCol)
    If anovaResult.Exists("error") Then
        ' UDF wrapper 必须返回 CVErr 而非字符串
        UDF_REGRESS_ANOVA = CVErr(xlErrValue)
    Else
        UDF_REGRESS_ANOVA = anovaResult("summary")
    End If
    Exit Function
ErrHandler:
    UDF_REGRESS_ANOVA = CVErr(xlErrValue)
End Function

Public Function UDF_REGRESS_PREDICT(ByVal data As Variant, ByVal factorCols As Variant, _
                            ByVal resultCol As Variant, ByVal newVals As Variant) As Variant
    Dim mdl As Object
    On Error GoTo ErrHandler
    Set mdl = LinearModelFit(data, ParseFactorCols(factorCols), resultCol)
    UDF_REGRESS_PREDICT = LinearModelPredict(mdl, newVals)
    Exit Function
ErrHandler:
    UDF_REGRESS_PREDICT = CVErr(xlErrValue)
End Function

Public Function UDF_REGRESS_SWEEP(ByVal data As Variant, ByVal factorCols As Variant, _
                          ByVal resultCol As Variant, ByVal sweepFactorIdx As Variant, _
                          ByVal fromVal As Variant, ByVal toVal As Variant, _
                          ByVal steps As Variant) As Variant
    On Error GoTo ErrHandler
    Dim mdl2 As Object: Set mdl2 = LinearModelFit(data, ParseFactorCols(factorCols), resultCol)
    UDF_REGRESS_SWEEP = FactorSweep(mdl2, sweepFactorIdx, fromVal, toVal, steps)
    Exit Function
ErrHandler:
    UDF_REGRESS_SWEEP = CVErr(xlErrValue)
End Function

Public Function UDF_REGRESS_OPTIMIZE(ByVal data As Variant, ByVal factorCols As Variant, _
                             ByVal resultCol As Variant, Optional ByVal goal As Variant = "max", _
                             Optional ByVal topN As Variant = 10, _
                             Optional ByVal nSteps As Variant = 10) As Variant
    On Error GoTo ErrHandler
    UDF_REGRESS_OPTIMIZE = OptimizeFactors(data, ParseFactorCols(factorCols), resultCol, goal, topN, nSteps)
    Exit Function
ErrHandler:
    UDF_REGRESS_OPTIMIZE = CVErr(xlErrValue)
End Function

'=====================================================================
' 使用示例
'=====================================================================
' ' --- 准备数据 ---
' Dim data As Range: Set data = Sheet1.Range("A1").CurrentRegion
' Dim factorCols() As Variant: factorCols = Array(1,2,3,4,5,6,7,8)
' Dim resultCol As Long: resultCol = 9
'
' ' --- 相关系数矩阵 ---
' Dim corr As Variant: corr = CorrelationMatrix(data)
'
' ' --- 因子重要性排序 ---
' Dim importance As Variant: importance = FactorImportance(data, factorCols, resultCol)
'
' ' --- 交互效应 ---
' Dim interact As Variant: interact = InteractionEffects(data, factorCols, resultCol)
'
' ' --- 单因素方差分析 ---
' Dim anova As Object: Set anova = ANOVAOneWay(data, 4, resultCol)
' anova("summary")       → "F(1,14) = 12.345, p = 0.0032, eta² = 0.469"
'
' ' --- 线性回归 ---
' Dim model As Object: Set model = LinearModelFit(data, factorCols, resultCol)
' model("formula")        → 公式字符串
' model("r_squared")      → R²
' model("adj_r_squared")  → 调整 R²
'
' ' --- 预测 ---
' Dim newVals As Variant: newVals = Array(2, 1.8, -0.5, "广东", "深圳", "南山", 1, 1)
' LinearModelPredict(model, newVals) → 预测值
'
' ' --- What-If 扫描 ---
' Dim sweep As Variant: sweep = FactorSweep(model, 1, 0#, 10#, 11)
'
' ' --- 最优因子组合 ---
' Dim best As Variant: best = OptimizeFactors(data, factorCols, resultCol, "max", 10, 10)

'=====================================================================
' Test_RegressUtils — 回归分析函数单元测试
'
' 在当前 VBA 项目中运行:   Call Test_RegressUtils