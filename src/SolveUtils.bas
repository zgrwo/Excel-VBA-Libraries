Option Explicit

'==============================================================================
' Module:       SolveUtils
' Purpose:      Process-parameter inversion: schema parsing, forward models
'               (linear/poly), cross-validation, bounded multi-start inversion,
'               reachability and report tables.
' Layer:        Analytics
' Dependencies: VBA-Core + LinearUtils (import these first)
' Public:       8 functions/subs
' Notes:        SOLVE.* v1 (auto/linear/poly). Ported from ExcelFormulaLabs SolveCore.
'==============================================================================

'=====================================================================
' SolveUtils.bas — 工艺参数反解 (SOLVE.*)
'
' 工作表函数 (UDF_SOLVE_*):
'   UDF_SOLVE_INVERSE   — 给定输出目标反推可调参数 (推荐表)
'   UDF_SOLVE_PREDICT   — 前向预测 (一行或多行完整参数)
'   UDF_SOLVE_QUALITY   — 各输出/候选模型的交叉验证质量表
'   UDF_SOLVE_EQUATION  — 前向方程 + 单可调线性闭式反解公式
'
' 依赖:
'   - LinearUtils.bas (QRDecomposition, MatrixTranspose, MatrixMultiply)
'   - VBA-Core (VariantKit, ArrayOps)
' 局限:
'   - v1 仅 auto/linear/poly; rate/rate_poly 与 SharedOutput 池化留待 v2
'   - 规模上限见模块常量 (历史 2000 行 / 请求 100 / 特征 30 / 可调 10 / 输出 10)
'=====================================================================

' ---- 错误码 (vbObjectError + 1601..1699) ----
Private Const ERR_SV_INVALID_INPUT As Long = vbObjectError + 1601
Private Const ERR_SV_NO_VARIABLE As Long = vbObjectError + 1602
Private Const ERR_SV_NO_OUTPUT As Long = vbObjectError + 1603
Private Const ERR_SV_BAD_MODEL As Long = vbObjectError + 1604
Private Const ERR_SV_TOO_FEW_ROWS As Long = vbObjectError + 1605
Private Const ERR_SV_POLY_LIMIT As Long = vbObjectError + 1606
Private Const ERR_SV_LIMIT As Long = vbObjectError + 1607
Private Const ERR_SV_NO_REQUEST As Long = vbObjectError + 1608
Private Const ERR_SV_BAD_BOUNDS As Long = vbObjectError + 1609
Private Const ERR_SV_NONFINITE As Long = vbObjectError + 1610
Private Const ERR_SV_DIV_ZERO As Long = vbObjectError + 1611
Private Const ERR_SV_CV As Long = vbObjectError + 1612
Private Const ERR_SV_NO_HISTORY As Long = vbObjectError + 1613

' ---- 规模上限 (port plan §6) ----
Private Const MAX_HISTORY_ROWS As Long = 2000
Private Const MAX_REQUEST_ROWS As Long = 100
Private Const MAX_FEATURE_COLS As Long = 30
Private Const MAX_VARIABLES As Long = 10
Private Const MAX_OUTPUTS As Long = 10
Private Const POLY_TERM_LIMIT As Long = 100
Private Const MIN_STARTS As Long = 1
Private Const MAX_STARTS_LIMIT As Long = 20
Private Const MAX_EVALS_PER_START As Long = 2000
Private Const REACH_SAMPLES As Long = 1000
Private Const POLY_RIDGE_LAMBDA As Double = 0.00001
Private Const TARGET_PRIORITY As Double = 1000000#
Private Const PROXIMITY_WEIGHT As Double = 0.02
Private Const STEP_FLOOR_FRACTION As Double = 0.000000001
Private Const REACH_TOL_FRACTION As Double = 0.000000001
Private Const ACHIEVED_SIGMA_TOL As Double = 0.000001
Private Const R2_TIE_TOL As Double = 0.000000001

Private Const MODEL_AUTO As String = "auto"
Private Const MODEL_LINEAR As String = "linear"
Private Const MODEL_POLY As String = "poly"
Private Const SCHEME_FIVE As String = "5折"
Private Const SCHEME_LOO As String = "LOO"
Private Const SCHEME_SKIP As String = "跳过"
Private Const STATUS_REACHABLE As String = "可达"
Private Const STATUS_UNREACHABLE As String = "不可达"

Private VK As VariantKit
Private AO As New ArrayOps

'=============================================================================
' 类型
'=============================================================================

Private Type TSvSchema
    ' 定长数组 + 计数成员 (VBA 不支持可靠地 ReDim UDT 成员数组)
    Headers(0 To 127) As String
    Incoming(0 To 63) As Long
    Variable(0 To 15) As Long
    FixedCols(0 To 63) As Long
    Output(0 To 15) As Long
    FeatureCols(0 To 63) As Long
    HistoryRows(0 To 2047) As Long
    RequestRows(0 To 255) As Long
    IncomingCount As Long
    VariableCount As Long
    FixedCount As Long
    OutputCount As Long
    FeatureCount As Long
    HistoryCount As Long
    RequestCount As Long
    ColCount As Long
End Type

Private Type TSvModel
    Kind As String
    BaseCount As Long
    TermCount As Long
    Intercept As Double
    Coef(0 To 100) As Double          ' 1-based, 1..TermCount
    Powers(0 To 100, 0 To 31) As Byte ' (1..TermCount, 1..BaseCount); 指数仅 0/1/2
    RankDeficient As Boolean          ' 增广 QR 出现零主元 → 部分系数不可辨识 (已置 0)
End Type

Private Type TSvRng
    Lo As Long
    Hi As Long
End Type

'=============================================================================
' 基础守卫 / 小工具
'=============================================================================

Private Sub SvEnsureCore()
    If VK Is Nothing Then Set VK = New VariantKit
End Sub

Private Function SvIsFinite(ByVal d As Double) As Boolean
    If d <> d Then Exit Function
    If d > 1.7976931348623157E+308 Then Exit Function
    If d < -1.7976931348623157E+308 Then Exit Function
    SvIsFinite = True
End Function

Private Function SvLimit(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If v < lo Then
        SvLimit = lo
    ElseIf v > hi Then
        SvLimit = hi
    Else
        SvLimit = v
    End If
End Function

' C# SolveCore.IsBlank: null/DBNull/ExcelEmpty/ExcelMissing/空白串
Private Function SvIsBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then SvIsBlank = True: Exit Function
    If VarType(v) = vbString Then
        SvIsBlank = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

Private Function SvIsNumericCell(ByVal v As Variant) As Boolean
    SvEnsureCore
    SvIsNumericCell = VK.IsNumericCell(v)
End Function

Private Function SvStartsWith(ByVal s As String, ByVal prefix As String, _
                              Optional ByVal ignoreCase As Boolean = False) As Boolean
    If Len(s) < Len(prefix) Then Exit Function
    If ignoreCase Then
        SvStartsWith = (StrComp(Left$(s, Len(prefix)), prefix, vbTextCompare) = 0)
    Else
        SvStartsWith = (Left$(s, Len(prefix)) = prefix)
    End If
End Function

Private Sub SvAppendLong(ByRef arr() As Long, ByRef cnt As Long, ByVal v As Long)
    If cnt = 0 Then
        ReDim arr(0 To 0)
    ElseIf cnt > UBound(arr) Then
        ReDim Preserve arr(0 To UBound(arr) * 2 + 1)
    End If
    arr(cnt) = v
    cnt = cnt + 1
End Sub

'=============================================================================
' Task 1.1 — 表解析与角色识别
'=============================================================================

Private Function SvClassifyHeader(ByVal header As String) As String
    ' 返回 "I"/"V"/"F"/"O"；未识别返回 ""
    If SvStartsWith(header, "Incoming", True) Or SvStartsWith(header, "来料") Then
        SvClassifyHeader = "I": Exit Function
    End If
    If SvStartsWith(header, "Variable", True) Or SvStartsWith(header, "可调") _
       Or SvStartsWith(header, "变量") Then
        SvClassifyHeader = "V": Exit Function
    End If
    If SvStartsWith(header, "Fixed", True) Or SvStartsWith(header, "固定") Then
        SvClassifyHeader = "F": Exit Function
    End If
    If SvStartsWith(header, "SharedOutput", True) Or SvStartsWith(header, "共享输出") Then
        SvClassifyHeader = "O": Exit Function   ' v1: 共享角色按普通输出处理
    End If
    If SvStartsWith(header, "Output", True) Or SvStartsWith(header, "输出") Then
        SvClassifyHeader = "O": Exit Function
    End If
End Function

' data: 2D Variant (Range.Value 或 NormalizeTo2D 结果)，任意 LBound
' requireHistory=False 用于 request 表解析 (只做表头角色映射, 不要求历史行)
Private Sub SvParseSchema(ByRef data As Variant, ByRef schema As TSvSchema, _
                          Optional ByVal requireHistory As Boolean = True)
    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long
    Dim rows As Long, cols As Long
    Dim c As Long, i As Long, r As Long, ci As Long
    Dim header As String, role As String
    Dim nInc As Long, nVar As Long, nFix As Long, nOut As Long, nFeat As Long
    Dim nHist As Long, nReq As Long
    Dim allVarNum As Boolean, anyVar As Boolean, allBlank As Boolean
    Dim v As Variant

    SvCoerceTableArg data, "data"
    If IsEmpty(data) Or Not IsArray(data) Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "Data table is empty."
    End If
    On Error Resume Next
    rLo = LBound(data, 1): rHi = UBound(data, 1)
    cLo = LBound(data, 2): cHi = UBound(data, 2)
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "Data must be a 2D range or array."
    End If
    On Error GoTo 0
    rows = rHi - rLo + 1: cols = cHi - cLo + 1
    If rows < 2 Or cols < 1 Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
            "Data table must contain a header row and at least one data row."
    End If

    nInc = 0: nVar = 0: nFix = 0: nOut = 0
    For c = 0 To cols - 1
        If c > 127 Then
            Err.Raise ERR_SV_LIMIT, "SolveUtils", "Too many columns: " & cols & " (limit 128)."
        End If
        header = Trim$(CStr(data(rLo, cLo + c)))
        schema.Headers(c) = header
        role = SvClassifyHeader(header)
        Select Case role
            Case "I"
                If nInc > UBound(schema.Incoming) Then
                    Err.Raise ERR_SV_LIMIT, "SolveUtils", "Too many incoming columns (hard limit " & (UBound(schema.Incoming) + 1) & ")."
                End If
                schema.Incoming(nInc) = c: nInc = nInc + 1
            Case "V"
                If nVar > UBound(schema.Variable) Then
                    Err.Raise ERR_SV_LIMIT, "SolveUtils", "Too many variable columns (hard limit " & (UBound(schema.Variable) + 1) & ")."
                End If
                schema.Variable(nVar) = c: nVar = nVar + 1
            Case "F"
                If nFix > UBound(schema.FixedCols) Then
                    Err.Raise ERR_SV_LIMIT, "SolveUtils", "Too many fixed columns (hard limit " & (UBound(schema.FixedCols) + 1) & ")."
                End If
                schema.FixedCols(nFix) = c: nFix = nFix + 1
            Case "O"
                If nOut > UBound(schema.Output) Then
                    Err.Raise ERR_SV_LIMIT, "SolveUtils", "Too many output columns (hard limit " & (UBound(schema.Output) + 1) & ")."
                End If
                schema.Output(nOut) = c: nOut = nOut + 1
        End Select
    Next c
    If nVar = 0 Then
        Err.Raise ERR_SV_NO_VARIABLE, "SolveUtils", _
            "No Variable column found. Name adjustable columns with a 'Variable*' (or '可调*'/'变量*') prefix."
    End If
    If nOut = 0 Then
        Err.Raise ERR_SV_NO_OUTPUT, "SolveUtils", _
            "No Output column found. Name result columns with an 'Output*' (or '输出*') prefix."
    End If
    If nInc + nVar + nFix > MAX_FEATURE_COLS Then
        Err.Raise ERR_SV_LIMIT, "SolveUtils", _
            "Too many feature columns: " & (nInc + nVar + nFix) & " (limit " & MAX_FEATURE_COLS & ")."
    End If
    If nVar > MAX_VARIABLES Then
        Err.Raise ERR_SV_LIMIT, "SolveUtils", _
            "Too many variable columns: " & nVar & " (limit " & MAX_VARIABLES & ")."
    End If
    If nOut > MAX_OUTPUTS Then
        Err.Raise ERR_SV_LIMIT, "SolveUtils", _
            "Too many output columns: " & nOut & " (limit " & MAX_OUTPUTS & ")."
    End If

    nFeat = nInc + nVar + nFix
    ci = 0
    For i = 0 To nInc - 1: schema.FeatureCols(ci) = schema.Incoming(i): ci = ci + 1: Next i
    For i = 0 To nVar - 1: schema.FeatureCols(ci) = schema.Variable(i): ci = ci + 1: Next i
    For i = 0 To nFix - 1: schema.FeatureCols(ci) = schema.FixedCols(i): ci = ci + 1: Next i
    schema.IncomingCount = nInc: schema.VariableCount = nVar
    schema.FixedCount = nFix: schema.OutputCount = nOut
    schema.FeatureCount = nFeat
    schema.ColCount = cols

    nHist = 0: nReq = 0
    For i = 1 To rows - 1          ' 0-based 数据行
        r = rLo + i
        allBlank = True
        For c = 0 To cols - 1
            If Not SvIsBlank(data(r, cLo + c)) Then allBlank = False: Exit For
        Next c
        If allBlank Then GoTo NextRow

        allVarNum = True: anyVar = False
        For ci = 0 To nVar - 1
            v = data(r, cLo + schema.Variable(ci))
            If SvIsBlank(v) Then
                allVarNum = False
            Else
                anyVar = True
                If Not SvIsNumericCell(v) Then allVarNum = False
            End If
        Next ci
        If allVarNum And anyVar Then
            If nHist > UBound(schema.HistoryRows) Then
                Err.Raise ERR_SV_LIMIT, "SolveUtils", "Too many history rows (hard limit " & (UBound(schema.HistoryRows) + 1) & ")."
            End If
            schema.HistoryRows(nHist) = i - 1: nHist = nHist + 1
            GoTo NextRow
        End If
        ' 请求行: 来料必须已知
        For ci = 0 To nInc - 1
            v = data(r, cLo + schema.Incoming(ci))
            If SvIsBlank(v) Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "Incoming column '" & schema.Headers(schema.Incoming(ci)) & "' is blank in request row " & (i + 1) & "."
            End If
            If Not SvIsNumericCell(v) Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "Incoming column '" & schema.Headers(schema.Incoming(ci)) & "' contains a non-numeric value in request row " & (i + 1) & "."
            End If
        Next ci
        If nReq > UBound(schema.RequestRows) Then
            Err.Raise ERR_SV_LIMIT, "SolveUtils", "Too many request rows (hard limit " & (UBound(schema.RequestRows) + 1) & ")."
        End If
        schema.RequestRows(nReq) = i - 1: nReq = nReq + 1
NextRow:
    Next i
    schema.HistoryCount = nHist
    schema.RequestCount = nReq
    If nHist > MAX_HISTORY_ROWS Then
        Err.Raise ERR_SV_LIMIT, "SolveUtils", _
            "Too many history rows: " & nHist & " (limit " & MAX_HISTORY_ROWS & ")."
    End If
    If nReq > MAX_REQUEST_ROWS Then
        Err.Raise ERR_SV_LIMIT, "SolveUtils", _
            "Too many request rows: " & nReq & " (limit " & MAX_REQUEST_ROWS & ")."
    End If
    If requireHistory Then
        If nHist = 0 Then
            Err.Raise ERR_SV_NO_HISTORY, "SolveUtils", "Data table contains no history rows."
        End If
    End If
End Sub

'=============================================================================
' Task 1.2 — 模型拟合与预测
'=============================================================================

Private Function SvNormalizeModel(ByVal model As Variant) As String
    Dim m As String
    If IsMissing(model) Or IsEmpty(model) Then m = MODEL_AUTO Else m = LCase$(Trim$(CStr(model)))
    Select Case m
        Case MODEL_AUTO, MODEL_LINEAR, MODEL_POLY: SvNormalizeModel = m
        Case Else
            Err.Raise ERR_SV_BAD_MODEL, "SolveUtils", _
                "Unknown model '" & CStr(model) & "'. Use ""auto"", ""linear"" or ""poly""."
    End Select
End Function

Private Function SvRequireFitModel(ByVal model As Variant) As String
    Dim m As String: m = SvNormalizeModel(model)
    If m = MODEL_AUTO Then
        Err.Raise ERR_SV_BAD_MODEL, "SolveUtils", _
            "FitModel requires an explicit model: ""linear"" or ""poly""."
    End If
    SvRequireFitModel = m
End Function

Private Function SvExpandedTermCount(ByVal baseCount As Long, ByVal m As String) As Long
    Select Case m
        Case MODEL_LINEAR: SvExpandedTermCount = baseCount
        Case MODEL_POLY: SvExpandedTermCount = baseCount + baseCount * (baseCount + 1) \ 2
        Case Else
            Err.Raise ERR_SV_BAD_MODEL, "SolveUtils", "Unknown model '" & m & "'."
    End Select
End Function

Private Sub SvBuildPowers(ByVal baseCount As Long, ByVal m As String, _
                          ByRef powers() As Byte, ByRef termCount As Long)
    Dim t As Long, j As Long, a As Long, b As Long
    termCount = SvExpandedTermCount(baseCount, m)
    ReDim powers(1 To termCount, 1 To baseCount)
    t = 1
    For j = 0 To baseCount - 1
        powers(t, j + 1) = 1: t = t + 1
    Next j
    If m = MODEL_POLY Then
        For j = 0 To baseCount - 1
            powers(t, j + 1) = 2: t = t + 1
        Next j
        For a = 0 To baseCount - 1
            For b = a + 1 To baseCount - 1
                powers(t, a + 1) = 1: powers(t, b + 1) = 1: t = t + 1
            Next b
        Next a
    End If
End Sub

' 1D 行向量 x (0-based, 长度 baseCount) 的项值
Private Function SvPowerProduct(ByRef x() As Double, ByRef powers() As Byte, _
                                ByVal term As Long, ByVal baseCount As Long) As Double
    Dim v As Double: v = 1#
    Dim j As Long, e As Long, p As Double, tt As Long
    For j = 1 To baseCount
        e = powers(term, j)
        If e = 1 Then
            v = v * x(j - 1)
        ElseIf e = 2 Then
            v = v * x(j - 1) * x(j - 1)
        ElseIf e > 0 Then
            p = 1#
            For tt = 1 To e: p = p * x(j - 1): Next tt
            v = v * p
        End If
    Next j
    SvPowerProduct = v
End Function

' 2D 矩阵 X (rowIdx 行, 0-based 列) 的项值
Private Function SvPowerProductM(ByRef X() As Double, ByVal rowIdx As Long, _
                                 ByRef powers() As Byte, ByVal term As Long, _
                                 ByVal baseCount As Long) As Double
    Dim v As Double: v = 1#
    Dim j As Long, e As Long, p As Double, tt As Long
    For j = 1 To baseCount
        e = powers(term, j)
        If e = 1 Then
            v = v * X(rowIdx, j - 1)
        ElseIf e = 2 Then
            v = v * X(rowIdx, j - 1) * X(rowIdx, j - 1)
        ElseIf e > 0 Then
            p = 1#
            For tt = 1 To e: p = p * X(rowIdx, j - 1): Next tt
            v = v * p
        End If
    Next j
    SvPowerProductM = v
End Function

' 列均值/样本标准差；溢出时按 max 归一化重算 (对齐 C# 审查 F1 回退)
Private Sub SvMoments(ByRef col() As Double, ByVal n As Long, ByRef mu As Double, ByRef sd As Double)
    Dim i As Long, d As Double
    Dim sum As Double, ss As Double, maxAbs As Double, a As Double
    Dim muN As Double, ssN As Double

    On Error GoTo Fallback
    sum = 0#: maxAbs = 0#
    For i = 0 To n - 1
        sum = sum + col(i)
        a = Abs(col(i)): If a > maxAbs Then maxAbs = a
    Next i
    mu = sum / n
    ss = 0#
    For i = 0 To n - 1
        d = col(i) - mu
        ss = ss + d * d
    Next i
    If SvIsFinite(ss) Then
        If n > 1 Then sd = Sqr(ss / (n - 1)) Else sd = 0#
        If SvIsFinite(sd) Then Exit Sub
    End If

Fallback:
    On Error GoTo 0
    If maxAbs = 0# Then
        mu = 0#: sd = 0#: Exit Sub
    End If
    muN = 0#
    For i = 0 To n - 1
        muN = muN + (col(i) / maxAbs - muN) / (i + 1)
    Next i
    ssN = 0#
    For i = 0 To n - 1
        d = col(i) / maxAbs - muN
        ssN = ssN + d * d
    Next i
    mu = muN * maxAbs
    If n > 1 Then sd = Sqr(ssN / (n - 1)) * maxAbs Else sd = 0#
    If Not SvIsFinite(mu) Then mu = 0#
    If Not SvIsFinite(sd) Then sd = 0#
End Sub

' 增广 QR 求解 [Z | 1] 的 OLS/岭回归 (λ 仅惩罚非截距项)
Private Sub SvSolveRidge(ByRef Z() As Double, ByVal n As Long, ByVal t As Long, _
                         ByRef y() As Double, ByVal lambda As Double, _
                         ByRef intercept As Double, ByRef coef() As Double, _
                         ByRef rankDeficient As Boolean)
    Dim m As Long: m = n + t
    Dim p As Long: p = t + 1
    Dim A() As Double, yA() As Double
    Dim Q() As Double, R() As Double
    Dim row As Long, col As Long, i As Long, j As Long
    Dim sq As Double
    Dim maxAbsR As Double, tol As Double
    Dim b() As Double, s As Double
    Dim rr0 As Long, rc0 As Long

    ReDim A(1 To m, 1 To p)
    ReDim yA(1 To m, 1 To 1)
    For i = 0 To n - 1
        For j = 0 To t - 1
            A(i + 1, j + 1) = Z(i, j)
        Next j
        A(i + 1, p) = 1#
        yA(i + 1, 1) = y(i)
    Next i
    If lambda > 0# Then
        sq = Sqr(lambda)
    Else
        sq = 0#
    End If
    For j = 0 To t - 1
        A(n + j + 1, j + 1) = sq
    Next j

    QRDecomposition A, Q, R, True
    ' R 为 maxK×maxK = p×p (m>=p)
    rr0 = LBound(R, 1): rc0 = LBound(R, 2)
    maxAbsR = 0#
    For i = 0 To p - 1
        For j = i To p - 1
            If Abs(R(rr0 + i, rc0 + j)) > maxAbsR Then maxAbsR = Abs(R(rr0 + i, rc0 + j))
        Next j
    Next i
    tol = maxAbsR * 2.22044604925031E-16 * CDbl(p)

    ' b = Q^T yA
    Dim Qt() As Double, Qty() As Double
    Qt = MatrixTranspose(Q)
    Qty = MatrixMultiply(Qt, yA)
    ReDim b(1 To p)
    rankDeficient = False
    For i = p - 1 To 0 Step -1
        s = Qty(LBound(Qty, 1) + i, LBound(Qty, 2))
        For j = i + 1 To p - 1
            s = s - R(rr0 + i, rc0 + j) * b(j + 1)
        Next j
        If R(rr0 + i, rc0 + i) = 0# Or (tol > 0# And Abs(R(rr0 + i, rc0 + i)) < tol) Then
            b(i + 1) = 0#
            rankDeficient = True
        Else
            b(i + 1) = s / R(rr0 + i, rc0 + i)
        End If
    Next i

    ReDim coef(0 To t - 1)
    For j = 0 To t - 1: coef(j) = b(j + 1): Next j
    intercept = b(p)
End Sub

' X: (0..n-1, 0..k-1); y: 0-based; model.Powers/Coef 为反标准化后的原始单位
Private Sub SvFitExpanded(ByRef X() As Double, ByVal n As Long, ByVal k As Long, _
                          ByRef y() As Double, ByVal kind As String, _
                          ByVal ridged As Boolean, ByRef model As TSvModel)
    Dim powers() As Byte, termCount As Long
    Dim t As Long, i As Long, j As Long
    Dim col() As Double
    Dim meanT() As Double, scaleT() As Double
    Dim active() As Long, nActive As Long
    Dim mu As Double, sd As Double, v As Double
    Dim Z() As Double
    Dim beta() As Double, b0 As Double
    Dim muY As Double, lambda As Double

    SvBuildPowers k, kind, powers, termCount
    If kind = MODEL_POLY And termCount > POLY_TERM_LIMIT Then
        Err.Raise ERR_SV_POLY_LIMIT, "SolveUtils", _
            "Poly expansion has " & termCount & " terms, which exceeds the limit of " & POLY_TERM_LIMIT & "."
    End If
    If n < termCount + 1 Then
        Err.Raise ERR_SV_TOO_FEW_ROWS, "SolveUtils", _
            "Not enough history rows for " & kind & " model: " & n & " rows for " & termCount & " terms (need at least " & (termCount + 1) & ")."
    End If

    model.Kind = kind
    model.BaseCount = k
    model.TermCount = termCount
    model.RankDeficient = False
    For i = 1 To termCount
        For j = 1 To k: model.Powers(i, j) = powers(i, j): Next j
    Next i
    ReDim meanT(1 To termCount), scaleT(1 To termCount)
    ReDim active(0 To termCount - 1)
    nActive = 0

    For t = 1 To termCount
        ReDim col(0 To n - 1)
        For i = 0 To n - 1
            v = SvPowerProductM(X, i, powers, t, k)
            If Not SvIsFinite(v) Then
                Err.Raise ERR_SV_NONFINITE, "SolveUtils", _
                    "Model term " & t & " is not representable in double precision; scale feature columns down or use model=""linear""."
            End If
            col(i) = v
        Next i
        SvMoments col, n, mu, sd
        meanT(t) = mu
        If sd > 0# Then
            scaleT(t) = sd
        Else
            scaleT(t) = 1#
        End If
        If sd > 0# Then
            active(nActive) = t: nActive = nActive + 1
        End If
        ' 复用 col 作为第一项存储的临时变量名无必要 — 下一轮重算; X 未变, PowerProduct 相同
        ' (为控制内存, 不缓存 raw 列; 第二轮构建 Z 时重算)
    Next t

    If nActive = 0 Then
        ' 仅截距模型
        muY = 0#
        For i = 0 To n - 1: muY = muY + (y(i) - muY) / (i + 1): Next i
        model.Intercept = muY
        For t = 1 To termCount: model.Coef(t) = 0#: Next t
        Exit Sub
    End If

    ReDim Z(0 To n - 1, 0 To nActive - 1)
    For i = 0 To n - 1
        For j = 0 To nActive - 1
            t = active(j)
            Z(i, j) = (SvPowerProductM(X, i, powers, t, k) - meanT(t)) / scaleT(t)
        Next j
    Next i

    If ridged Then
        lambda = POLY_RIDGE_LAMBDA
    Else
        lambda = 0#
    End If
    SvSolveRidge Z, n, nActive, y, lambda, b0, beta, model.RankDeficient

    ' 反标准化: coef[t] = β_a/scale[t]; intercept = b0 − Σ β_a·mean[t]/scale[t]
    model.Intercept = b0
    For t = 1 To termCount: model.Coef(t) = 0#: Next t
    For j = 0 To nActive - 1
        t = active(j)
        model.Coef(t) = beta(j) / scaleT(t)
        model.Intercept = model.Intercept - beta(j) * meanT(t) / scaleT(t)
    Next j
End Sub

Private Function SvPredict(ByRef model As TSvModel, ByRef x() As Double) As Double
    Dim t As Long
    Dim v As Double: v = model.Intercept
    For t = 1 To model.TermCount
        If model.Coef(t) <> 0# Then
            v = v + model.Coef(t) * SvPowerProduct(x, model.Powers, t, model.BaseCount)
        End If
    Next t
    SvPredict = v
End Function

'=============================================================================
' Task 1.3 — XorShift64* 确定性 RNG (与 C# XorShift64 同参数: 12/25/27 + 乘子)
'=============================================================================

' 32 位无符号 <-> 有符号 Long 视图
Private Function SvU32ToD(ByVal v As Long) As Double
    If v < 0 Then SvU32ToD = CDbl(v) + 4294967296# Else SvU32ToD = CDbl(v)
End Function

Private Function SvDToI32(ByVal d As Double) As Long
    d = d - Int(d / 4294967296#) * 4294967296#
    If d < 0# Then d = d + 4294967296#
    If d >= 2147483648# Then d = d - 4294967296#
    SvDToI32 = CLng(d)
End Function

' (v << n) mod 2^32
Private Function SvShl32(ByVal v As Long, ByVal n As Long) As Long
    Dim d As Double, lowPart As Double, m As Long
    If n <= 0 Then SvShl32 = v: Exit Function
    If n >= 32 Then SvShl32 = 0: Exit Function
    m = 32 - n
    d = SvU32ToD(v)
    lowPart = d - Int(d / 2# ^ m) * 2# ^ m
    SvShl32 = SvDToI32(lowPart * 2# ^ n)
End Function

' 逻辑右移 (无符号语义)
Private Function SvShr32(ByVal v As Long, ByVal n As Long) As Long
    If n <= 0 Then SvShr32 = v: Exit Function
    If n >= 32 Then SvShr32 = 0: Exit Function
    SvShr32 = CLng(Int(SvU32ToD(v) / 2# ^ n))
End Function

' 32x32 -> 64 位乘积 (双 limb)
Private Sub SvMulU32(ByVal a As Long, ByVal b As Long, ByRef lo As Long, ByRef hi As Long)
    Dim ad As Double, bd As Double
    Dim alo As Double, ahi As Double, blo As Double, bhi As Double
    Dim p00 As Double, p01 As Double, p11 As Double
    Dim lim0 As Double, carry As Double

    ad = SvU32ToD(a): bd = SvU32ToD(b)
    alo = ad - Int(ad / 65536#) * 65536#: ahi = Int(ad / 65536#)
    blo = bd - Int(bd / 65536#) * 65536#: bhi = Int(bd / 65536#)
    p00 = alo * blo
    p01 = alo * bhi + ahi * blo
    p11 = ahi * bhi
    lim0 = p00 + (p01 - Int(p01 / 65536#) * 65536#) * 65536#
    carry = Int(p01 / 65536#) + p11 + Int(lim0 / 4294967296#)
    lo = SvDToI32(lim0)
    hi = SvDToI32(carry)
End Sub

' (aLo:aHi) * (bLo:bHi) mod 2^64
Private Sub SvMul64(ByVal aLo As Long, ByVal aHi As Long, ByVal bLo As Long, ByVal bHi As Long, _
                    ByRef outLo As Long, ByRef outHi As Long)
    Dim p1lo As Long, p1hi As Long
    Dim p2lo As Long, p2hi As Long, p3lo As Long, p3hi As Long

    SvMulU32 aLo, bLo, p1lo, p1hi
    SvMulU32 aLo, bHi, p2lo, p2hi
    SvMulU32 aHi, bLo, p3lo, p3hi
    outLo = p1lo
    outHi = SvDToI32(SvU32ToD(p1hi) + SvU32ToD(p2lo) + SvU32ToD(p3lo))
End Sub

Private Sub SvShr64(ByRef lo As Long, ByRef hi As Long, ByVal n As Long)
    Dim nlo As Long, nhi As Long
    nlo = SvShr32(lo, n) Or SvShl32(hi, 32 - n)
    nhi = SvShr32(hi, n)
    lo = nlo: hi = nhi
End Sub

Private Sub SvShl64(ByRef lo As Long, ByRef hi As Long, ByVal n As Long)
    Dim nlo As Long, nhi As Long
    nlo = SvShl32(lo, n)
    nhi = SvShl32(hi, n) Or SvShr32(lo, 32 - n)
    lo = nlo: hi = nhi
End Sub

Private Sub SvRngInit(ByRef rng As TSvRng, ByVal seed As Long)
    If seed = 0 Then
        rng.Lo = &H7F4A7C15
        rng.Hi = &H9E3779B9
    Else
        rng.Lo = seed
        If seed < 0 Then rng.Hi = -1 Else rng.Hi = 0
    End If
End Sub

' xorshift64*: state = xorshift(x); 返回 x * 0x2545F4914F6CDD1D
Private Sub SvRngNext(ByRef rng As TSvRng, ByRef outLo As Long, ByRef outHi As Long)
    Dim xlo As Long, xhi As Long
    Dim tlo As Long, thi As Long

    xlo = rng.Lo: xhi = rng.Hi
    tlo = xlo: thi = xhi
    SvShr64 tlo, thi, 12
    xlo = xlo Xor tlo: xhi = xhi Xor thi
    tlo = xlo: thi = xhi
    SvShl64 tlo, thi, 25
    xlo = xlo Xor tlo: xhi = xhi Xor thi
    tlo = xlo: thi = xhi
    SvShr64 tlo, thi, 27
    xlo = xlo Xor tlo: xhi = xhi Xor thi
    rng.Lo = xlo: rng.Hi = xhi
    SvMul64 xlo, xhi, &H4F6CDD1D, &H2545F491, outLo, outHi
End Sub

' (Next() >> 11) * 2^-53 — 与 C# NextUniform 一致
Private Function SvRngUniform(ByRef rng As TSvRng) As Double
    Dim lo As Long, hi As Long
    SvRngNext rng, lo, hi
    SvRngUniform = (SvU32ToD(hi) * 2097152# + CDbl(SvShr32(lo, 11))) / 9007199254740992#
End Function

' (long)(Next() % bound) — 与 C# NextLong 一致
Private Function SvRngIndex(ByRef rng As TSvRng, ByVal bound As Long) As Long
    Dim lo As Long, hi As Long
    Dim r As Double, bd As Double
    SvRngNext rng, lo, hi
    bd = CDbl(bound)
    r = SvU32ToD(hi)
    r = r - Int(r / bd) * bd
    r = r * 4294967296# + SvU32ToD(lo)
    r = r - Int(r / bd) * bd
    SvRngIndex = CLng(r)
End Function

'=============================================================================
' Task 1.3 — 交叉验证与 auto 门控
'=============================================================================

Private Function SvSampleStdDev(ByRef v() As Double, ByVal n As Long) As Double
    Dim i As Long, m As Double, ss As Double, d As Double
    m = 0#
    For i = 0 To n - 1: m = m + (v(i) - m) / (i + 1): Next i
    ss = 0#
    For i = 0 To n - 1
        d = v(i) - m
        ss = ss + d * d
    Next i
    If n > 1 And SvIsFinite(ss) Then SvSampleStdDev = Sqr(ss / (n - 1)) Else SvSampleStdDev = 0#
End Function

Private Sub SvSubsetX(ByRef X() As Double, ByVal k As Long, ByRef idx() As Long, _
                      ByVal lo As Long, ByVal hi As Long, ByRef out() As Double)
    Dim cnt As Long: cnt = hi - lo + 1
    Dim i As Long, j As Long
    ReDim out(0 To cnt - 1, 0 To k - 1)
    For i = 0 To cnt - 1
        For j = 0 To k - 1
            out(i, j) = X(idx(lo + i), j)
        Next j
    Next i
End Sub

Private Sub SvSubsetY(ByRef y() As Double, ByRef idx() As Long, _
                      ByVal lo As Long, ByVal hi As Long, ByRef out() As Double)
    Dim cnt As Long: cnt = hi - lo + 1
    Dim i As Long
    ReDim out(0 To cnt - 1)
    For i = 0 To cnt - 1: out(i) = y(idx(lo + i)): Next i
End Sub

Private Sub SvCrossValidate(ByRef X() As Double, ByVal n As Long, ByVal k As Long, _
                            ByRef y() As Double, ByVal modelName As String, ByVal seed As Long, _
                            ByRef scheme As String, ByRef r2 As Double, ByRef mae As Double)
    If n < 5 Then
        Err.Raise ERR_SV_CV, "SolveUtils", "Cross-validation needs at least 5 history rows (got " & n & ")."
    End If
    Dim order() As Long, i As Long, j As Long, tmp As Long
    Dim st As TSvRng
    SvRngInit st, seed
    Dim fiveFold As Boolean: fiveFold = (n >= 20)
    Dim folds As Long: If fiveFold Then folds = 5 Else folds = n
    Dim pred() As Double
    Dim f As Long, t As Long
    Dim trainIdx() As Long, testIdx() As Long
    Dim nTrain As Long, nTest As Long
    Dim Xt() As Double, yt() As Double, Xv() As Double
    Dim m As TSvModel, p As Double

    ReDim order(0 To n - 1)
    For i = 0 To n - 1: order(i) = i: Next i
    For i = n - 1 To 1 Step -1
        j = SvRngIndex(st, i + 1)
        tmp = order(i): order(i) = order(j): order(j) = tmp
    Next i

    ReDim pred(0 To n - 1)
    For f = 0 To folds - 1
        ReDim trainIdx(0 To n - 1): ReDim testIdx(0 To n - 1)
        nTrain = 0: nTest = 0
        For i = 0 To n - 1
            If (fiveFold And (i Mod 5 = f)) Or ((Not fiveFold) And i = f) Then
                testIdx(nTest) = order(i): nTest = nTest + 1
            Else
                trainIdx(nTrain) = order(i): nTrain = nTrain + 1
            End If
        Next i
        SvSubsetX X, k, trainIdx, 0, nTrain - 1, Xt
        SvSubsetY y, trainIdx, 0, nTrain - 1, yt
        SvSubsetX X, k, testIdx, 0, nTest - 1, Xv
        SvFitExpanded Xt, nTrain, k, yt, modelName, (modelName = MODEL_POLY), m
        Dim xrow() As Double
        For t = 0 To nTest - 1
            xrow = SvRow(Xv, t, k)
            p = SvPredict(m, xrow)
            If Not SvIsFinite(p) Then
                Err.Raise ERR_SV_CV, "SolveUtils", _
                    "Cross-validation produced a non-finite prediction; model is numerically unstable."
            End If
            pred(testIdx(t)) = p
        Next t
    Next f

    Dim yMean As Double, sse As Double, tss As Double
    yMean = 0#
    For i = 0 To n - 1: yMean = yMean + (y(i) - yMean) / (i + 1): Next i
    sse = 0#: tss = 0#: mae = 0#
    For i = 0 To n - 1
        sse = sse + (pred(i) - y(i)) ^ 2
        mae = mae + Abs(pred(i) - y(i))
        tss = tss + (y(i) - yMean) ^ 2
    Next i
    If Not SvIsFinite(sse) Or Not SvIsFinite(tss) Or Not SvIsFinite(mae) Then
        Err.Raise ERR_SV_CV, "SolveUtils", "Cross-validation statistics are numerically unstable."
    End If
    If tss = 0# Then
        Err.Raise ERR_SV_CV, "SolveUtils", "Cannot cross-validate: constant response variable y."
    End If
    If fiveFold Then
        scheme = SCHEME_FIVE
    Else
        scheme = SCHEME_LOO
    End If
    r2 = 1# - sse / tss
    mae = mae / n
End Sub

' Xv 的第 t 行视图 (返回新数组，简化调用)
Private Function SvRow(ByRef Xv() As Double, ByVal rowIdx As Long, ByVal k As Long) As Double()
    Dim out() As Double
    Dim j As Long
    ReDim out(0 To k - 1)
    For j = 0 To k - 1: out(j) = Xv(rowIdx, j): Next j
    SvRow = out
End Function

Private Function SvPolyExcludedBySample(ByVal n As Long, ByVal k As Long) As Boolean
    Dim terms As Long: terms = SvExpandedTermCount(k, MODEL_POLY)
    If terms > POLY_TERM_LIMIT Then SvPolyExcludedBySample = True: Exit Function
    Dim minTrain As Long
    If n >= 20 Then minTrain = n - (n + 4) \ 5 Else minTrain = n - 1
    SvPolyExcludedBySample = (minTrain < terms + 2)
End Function

Private Sub SvFitAuto(ByRef X() As Double, ByVal n As Long, ByVal k As Long, _
                      ByRef y() As Double, ByVal seed As Long, _
                      ByRef model As TSvModel, ByRef chosen As String, _
                      ByRef scheme As String, ByRef r2 As Double, ByRef mae As Double, _
                      ByRef polySkipped As Boolean)
    Dim linScheme As String, linR2 As Double, linMae As Double
    Dim polyScheme As String, polyR2 As Double, polyMae As Double
    Dim linSkipped As Boolean, firstErrNum As Long, firstErrDesc As String
    Dim ok As Boolean

    On Error Resume Next
    Err.Clear
    SvCrossValidate X, n, k, y, MODEL_LINEAR, seed, linScheme, linR2, linMae
    If Err.Number <> 0 Then
        linSkipped = True
        firstErrNum = Err.Number: firstErrDesc = Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    polySkipped = SvPolyExcludedBySample(n, k)
    If Not polySkipped Then
        On Error Resume Next
        Err.Clear
        SvCrossValidate X, n, k, y, MODEL_POLY, seed, polyScheme, polyR2, polyMae
        If Err.Number <> 0 Then
            polySkipped = True
            If firstErrNum = 0 Then firstErrNum = Err.Number: firstErrDesc = Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If

    If linSkipped And polySkipped Then
        If firstErrNum <> 0 Then
            Err.Raise firstErrNum, "SolveUtils", firstErrDesc
        Else
            Err.Raise ERR_SV_CV, "SolveUtils", "No model candidate could be cross-validated."
        End If
    End If

    If Not linSkipped Then
        chosen = MODEL_LINEAR: r2 = linR2: mae = linMae: scheme = linScheme
    Else
        chosen = MODEL_POLY: r2 = polyR2: mae = polyMae: scheme = polyScheme
    End If
    If Not polySkipped Then
        If polyR2 > r2 + R2_TIE_TOL Then
            chosen = MODEL_POLY: r2 = polyR2: mae = polyMae: scheme = polyScheme
        End If
    End If

    SvFitExpanded X, n, k, y, chosen, (chosen = MODEL_POLY), model
End Sub

'=============================================================================
' Task 1.4 — 反解优化器 + 可达性 (v1: linear/poly，无 rate/shared)
'=============================================================================

' 中位数 (排序副本; 偶数取中间两数均值)
Private Function SvMedian(ByRef v() As Double, ByVal n As Long) As Double
    Dim tmp() As Variant, i As Long
    ReDim tmp(0 To n - 1)
    For i = 0 To n - 1: tmp(i) = v(i): Next i
    AO.Sort tmp, True, "auto"
    If n Mod 2 = 1 Then
        SvMedian = CDbl(tmp((n - 1) \ 2))
    Else
        SvMedian = 0.5 * (CDbl(tmp(n \ 2 - 1)) + CDbl(tmp(n \ 2)))
    End If
End Function

' 目标函数: TARGET_PRIORITY·Σ_j((ŷ_j−y*_j)/s_j)² + PROXIMITY_WEIGHT·Σ_c((u_c−med_c)/range_c)²
' tgtMask 标记被瞄准的输出 (空目标由上层以掩码表达, 不依赖 NaN)
Private Function SvObjective(ByRef models() As TSvModel, ByRef target() As Double, _
                             ByRef tgtMask() As Boolean, ByRef scales() As Double, _
                             ByRef medians() As Double, ByRef bounds() As Double, _
                             ByRef variableCols() As Long, ByRef u() As Double, _
                             ByRef feature() As Double, ByVal outCount As Long) As Double
    Dim c As Long, j As Long
    Dim fTarget As Double, proximity As Double
    Dim pr As Double, dv As Double, rng As Double, z As Double

    For c = 0 To UBound(variableCols)
        feature(variableCols(c)) = u(c)
    Next c
    fTarget = 0#
    For j = 0 To outCount - 1
        If tgtMask(j) Then
            pr = SvPredict(models(j), feature)
            If Not SvIsFinite(pr) Then SvObjective = 1E+308: Exit Function
            dv = (pr - target(j)) / scales(j)
            fTarget = fTarget + dv * dv
            If Not SvIsFinite(fTarget) Then SvObjective = 1E+308: Exit Function
        End If
    Next j
    proximity = 0#
    For c = 0 To UBound(variableCols)
        rng = bounds(c, 1) - bounds(c, 0)
        If rng > 0# Then
            z = (u(c) - medians(variableCols(c))) / rng
            proximity = proximity + z * z
        End If
    Next c
    SvObjective = TARGET_PRIORITY * fTarget + PROXIMITY_WEIGHT * proximity
End Function

' 坐标轮换模式搜索 + 步长折半 (clamp 回边界); 每起点评估上限 MAX_EVALS_PER_START
' feature 为调用方的特征缓冲 (已含来料/固定请求值), 搜索只改写可调位置
Private Function SvLocalSearch(ByRef models() As TSvModel, ByRef target() As Double, _
                               ByRef tgtMask() As Boolean, ByRef scales() As Double, _
                               ByRef medians() As Double, ByRef bounds() As Double, _
                               ByRef variableCols() As Long, ByRef start() As Double, _
                               ByVal outCount As Long, ByRef feature() As Double) As Double()
    Dim v As Long: v = UBound(start) + 1
    Dim u() As Double, trial() As Double, step() As Double
    Dim c As Long, dir As Long, i As Long
    Dim cand As Double, e As Double, best As Double
    Dim evaluations As Long, improved As Boolean, anyLarge As Boolean
    Dim rng As Double

    ReDim u(0 To v - 1): ReDim trial(0 To v - 1): ReDim step(0 To v - 1)
    For c = 0 To v - 1
        u(c) = start(c)
        step(c) = (bounds(c, 1) - bounds(c, 0)) * 0.25
    Next c
    best = SvObjective(models, target, tgtMask, scales, medians, bounds, variableCols, u, feature, outCount)
    evaluations = 1
    Do
        improved = False
        For c = 0 To v - 1
            If step(c) > 0# Then
                For dir = 1 To -1 Step -2
                    cand = SvLimit(u(c) + dir * step(c), bounds(c, 0), bounds(c, 1))
                    If cand <> u(c) Then
                        For i = 0 To v - 1: trial(i) = u(i): Next i
                        trial(c) = cand
                        e = SvObjective(models, target, tgtMask, scales, medians, bounds, _
                                        variableCols, trial, feature, outCount)
                        evaluations = evaluations + 1
                        If e < best Then
                            For i = 0 To v - 1: u(i) = trial(i): Next i
                            best = e
                            improved = True
                        End If
                        If evaluations >= MAX_EVALS_PER_START Then
                            SvLocalSearch = u
                            Exit Function
                        End If
                    End If
                Next dir
            End If
        Next c
        If Not improved Then
            anyLarge = False
            For c = 0 To v - 1
                rng = bounds(c, 1) - bounds(c, 0)
                If rng > 0# And step(c) > STEP_FLOOR_FRACTION * rng Then
                    step(c) = step(c) * 0.5
                    anyLarge = True
                End If
            Next c
            If Not anyLarge Then Exit Do
        End If
    Loop
    SvLocalSearch = u
End Function

' 多起点反解核心。
' X (0..n-1,0..k-1) / Y (0..n-1,0..outCount-1) / variableCols 为特征位置 (0-based)
' requests (0..r-1,0..k-1) 已由装配层补齐数值 (空白初值以历史中位数预填)
' targets (0..r-1,0..outCount-1) + tgtMask 标记瞄准输出
' result (0..r-1, 0..v+outCount+1) = [建议值…, 预测…, 最大偏差σ, 状态(0=可达/1=不可达)]
Private Sub SvSolveInverseCore(ByRef X() As Double, ByVal n As Long, _
                               ByRef Y() As Double, ByVal outCount As Long, _
                               ByRef variableCols() As Long, _
                               ByRef requests() As Double, ByVal r As Long, _
                               ByRef targets() As Double, ByRef tgtMask() As Boolean, _
                               ByRef bounds() As Double, ByRef models() As TSvModel, _
                               ByVal seed As Long, ByVal maxStarts As Long, _
                               ByRef result() As Double)
    Dim k As Long: k = models(0).BaseCount
    Dim v As Long: v = UBound(variableCols) + 1
    Dim i As Long, j As Long, c As Long, q As Long, s As Long, st As Long
    Dim scales() As Double, medians() As Double
    Dim col() As Double, yj() As Double
    Dim feature() As Double, sampleU() As Double
    Dim target() As Double, tgtMaskRow() As Boolean
    Dim minOut() As Double, maxOut() As Double, finiteCnt() As Long
    Dim u0() As Double, bestU() As Double, found() As Double
    Dim bestE As Double, e As Double, pr As Double
    Dim lo As Double, hi As Double, init As Double, range As Double
    Dim magnitude As Double, tol As Double, dev As Double
    Dim inRange As Boolean, achieved As Boolean, allReachable As Boolean
    Dim maxDeviation As Double
    Dim haveBest As Boolean
    Dim rng As TSvRng

    ' 每输出偏差尺度: 历史样本 sd (0 → 1)
    ReDim scales(0 To outCount - 1)
    For j = 0 To outCount - 1
        ReDim yj(0 To n - 1)
        For i = 0 To n - 1: yj(i) = Y(i, j): Next i
        scales(j) = SvSampleStdDev(yj, n)
        If scales(j) <= 0# Then scales(j) = 1#
    Next j
    ' 每特征列中位数 (按数据列位置)
    ReDim medians(0 To k - 1)
    For c = 0 To k - 1
        ReDim col(0 To n - 1)
        For i = 0 To n - 1: col(i) = X(i, c): Next i
        medians(c) = SvMedian(col, n)
    Next c

    ReDim result(0 To r - 1, 0 To v + outCount + 1)
    ReDim feature(0 To k - 1)
    ReDim sampleU(0 To v - 1)
    ReDim target(0 To outCount - 1)
    ReDim tgtMaskRow(0 To outCount - 1)
    ReDim minOut(0 To outCount - 1)
    ReDim maxOut(0 To outCount - 1)
    ReDim finiteCnt(0 To outCount - 1)
    SvRngInit rng, seed

    For q = 0 To r - 1
        For c = 0 To k - 1: feature(c) = requests(q, c): Next c
        For j = 0 To outCount - 1
            target(j) = targets(q, j)
            tgtMaskRow(j) = tgtMask(q, j)
            minOut(j) = 0#: maxOut(j) = 0#: finiteCnt(j) = 0
        Next j

        ' 可达性: REACH_SAMPLES 个均匀样本 (前两个取端点)
        For s = 0 To REACH_SAMPLES - 1
            For c = 0 To v - 1
                lo = bounds(c, 0): hi = bounds(c, 1)
                If s = 0 Then
                    sampleU(c) = lo
                ElseIf s = 1 Then
                    sampleU(c) = hi
                Else
                    sampleU(c) = lo + (hi - lo) * SvRngUniform(rng)
                End If
            Next c
            For c = 0 To v - 1: feature(variableCols(c)) = sampleU(c): Next c
            For j = 0 To outCount - 1
                pr = SvPredict(models(j), feature)
                If SvIsFinite(pr) Then
                    If finiteCnt(j) = 0 Then
                        minOut(j) = pr: maxOut(j) = pr
                    Else
                        If pr < minOut(j) Then minOut(j) = pr
                        If pr > maxOut(j) Then maxOut(j) = pr
                    End If
                    finiteCnt(j) = finiteCnt(j) + 1
                End If
            Next j
        Next s
        For j = 0 To outCount - 1
            If finiteCnt(j) = 0 Then
                Err.Raise ERR_SV_NONFINITE, "SolveUtils", _
                    "Reachability sampling produced no finite predictions for output " & (j + 1) & "."
            End If
        Next j

        ' 多起点模式搜索
        bestE = 1E+308
        haveBest = False
        For st = 0 To maxStarts - 1
            ReDim u0(0 To v - 1)
            For c = 0 To v - 1
                lo = bounds(c, 0): hi = bounds(c, 1)
                If st = 0 Then
                    init = requests(q, variableCols(c))
                    u0(c) = SvLimit(init, lo, hi)
                Else
                    u0(c) = lo + (hi - lo) * SvRngUniform(rng)
                End If
            Next c
            found = SvLocalSearch(models, target, tgtMaskRow, scales, medians, bounds, _
                                  variableCols, u0, outCount, feature)
            e = SvObjective(models, target, tgtMaskRow, scales, medians, bounds, _
                            variableCols, found, feature, outCount)
            If e < bestE Then
                bestE = e
                bestU = found
                haveBest = True
            End If
        Next st
        If Not haveBest Then
            Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "Optimization failed to evaluate any start point."
        End If

        For c = 0 To v - 1: feature(variableCols(c)) = bestU(c): Next c
        maxDeviation = 0#
        allReachable = True
        For j = 0 To outCount - 1
            pr = SvPredict(models(j), feature)
            If Not SvIsFinite(pr) Then
                Err.Raise ERR_SV_NONFINITE, "SolveUtils", "Optimization produced a non-finite prediction."
            End If
            If tgtMaskRow(j) Then
                dev = Abs(pr - target(j)) / scales(j)
                If dev > maxDeviation Then maxDeviation = dev
                magnitude = Abs(minOut(j))
                If Abs(maxOut(j)) > magnitude Then magnitude = Abs(maxOut(j))
                If Abs(maxOut(j) - minOut(j)) > magnitude Then magnitude = Abs(maxOut(j) - minOut(j))
                If magnitude < 1E-300 Then magnitude = 1E-300
                tol = REACH_TOL_FRACTION * magnitude
                inRange = (target(j) >= minOut(j) - tol) And (target(j) <= maxOut(j) + tol)
                achieved = (dev <= ACHIEVED_SIGMA_TOL)
                If (Not inRange) And (Not achieved) Then allReachable = False
            End If
        Next j
        For c = 0 To v - 1: result(q, c) = bestU(c): Next c
        For j = 0 To outCount - 1: result(q, v + j) = SvPredict(models(j), feature): Next j
        result(q, v + outCount) = maxDeviation
        If allReachable Then
            result(q, v + outCount + 1) = 0#
        Else
            result(q, v + outCount + 1) = 1#
        End If
    Next q
End Sub



'=============================================================================
' Task 1.4b — 表装配层 (schema → 历史/请求/边界 → 核心 → 报告表)
'=============================================================================

Private Function SvFeaturePos(ByRef schema As TSvSchema, ByVal dataCol As Long) As Long
    Dim c As Long
    For c = 0 To schema.FeatureCount - 1
        If schema.FeatureCols(c) = dataCol Then SvFeaturePos = c: Exit Function
    Next c
    SvFeaturePos = -1
End Function

Private Function SvIsIncoming(ByRef schema As TSvSchema, ByVal dataCol As Long) As Boolean
    Dim i As Long
    For i = 0 To schema.IncomingCount - 1
        If schema.Incoming(i) = dataCol Then SvIsIncoming = True: Exit Function
    Next i
End Function

Private Function SvIsFixed(ByRef schema As TSvSchema, ByVal dataCol As Long) As Boolean
    Dim i As Long
    For i = 0 To schema.FixedCount - 1
        If schema.FixedCols(i) = dataCol Then SvIsFixed = True: Exit Function
    Next i
End Function

Private Function SvToDoubleCell(ByVal v As Variant) As Double
    Dim d As Double
    If SvIsBlank(v) Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "Expected a numeric value, got a blank."
    End If
    If Not SvIsNumericCell(v) Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "Expected a numeric value, got '" & CStr(v) & "'."
    End If
    d = CDbl(v)
    If Not SvIsFinite(d) Then
        Err.Raise ERR_SV_NONFINITE, "SolveUtils", "Numeric value is not finite."
    End If
    SvToDoubleCell = d
End Function

Private Sub SvFeatureMedians(ByRef X() As Double, ByVal n As Long, ByVal k As Long, ByRef med() As Double)
    Dim c As Long, i As Long
    Dim col() As Double
    ReDim med(0 To k - 1)
    For c = 0 To k - 1
        ReDim col(0 To n - 1)
        For i = 0 To n - 1: col(i) = X(i, c): Next i
        med(c) = SvMedian(col, n)
    Next c
End Sub

' 历史矩阵 X (n×k) / Y (n×outCount)；固定列空白以列中位数填补
Private Sub SvBuildHistory(ByRef data As Variant, ByRef schema As TSvSchema, _
                           ByRef X() As Double, ByRef Y() As Double)
    Dim n As Long: n = schema.HistoryCount
    Dim k As Long: k = schema.FeatureCount
    Dim outCount As Long: outCount = schema.OutputCount
    Dim rLo As Long, cLo As Long
    Dim i As Long, c As Long, j As Long, dataCol As Long
    Dim arrayRow As Long
    Dim cell As Variant, d As Double
    Dim miss() As Boolean
    Dim anyMiss As Boolean, cnt As Long
    Dim colVals() As Double, med As Double

    rLo = LBound(data, 1): cLo = LBound(data, 2)
    ReDim X(0 To n - 1, 0 To k - 1)
    ReDim Y(0 To n - 1, 0 To outCount - 1)
    ReDim miss(0 To n - 1, 0 To k - 1)

    For i = 0 To n - 1
        arrayRow = rLo + schema.HistoryRows(i) + 1
        For c = 0 To k - 1
            dataCol = schema.FeatureCols(c)
            cell = data(arrayRow, cLo + dataCol)
            If SvIsBlank(cell) Then
                If SvIsFixed(schema, dataCol) Then
                    miss(i, c) = True
                Else
                    Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                        "History row " & (schema.HistoryRows(i) + 2) & " is missing a value in column '" & schema.Headers(dataCol) & "'."
                End If
            Else
                d = SvToDoubleCell(cell)
                X(i, c) = d
            End If
        Next c
        For j = 0 To outCount - 1
            cell = data(arrayRow, cLo + schema.Output(j))
            If SvIsBlank(cell) Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "History row " & (schema.HistoryRows(i) + 2) & " has a non-numeric output in column '" & schema.Headers(schema.Output(j)) & "'."
            End If
            d = SvToDoubleCell(cell)
            Y(i, j) = d
        Next j
    Next i

    ' 固定列空白 → 列中位数
    For c = 0 To k - 1
        anyMiss = False
        For i = 0 To n - 1
            If miss(i, c) Then anyMiss = True: Exit For
        Next i
        If anyMiss Then
            ReDim colVals(0 To n - 1)
            cnt = 0
            For i = 0 To n - 1
                If Not miss(i, c) Then colVals(cnt) = X(i, c): cnt = cnt + 1
            Next i
            If cnt = 0 Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "Fixed column '" & schema.Headers(schema.FeatureCols(c)) & "' has no numeric history value to impute."
            End If
            ReDim Preserve colVals(0 To cnt - 1)
            med = SvMedian(colVals, cnt)
            For i = 0 To n - 1
                If miss(i, c) Then X(i, c) = med
            Next i
        End If
    Next c
End Sub

' 请求行特征向量：来料缺失/非数值 → Err；固定缺失/非数值 → 历史中位数；
' 可调缺失/非数值 → 历史中位数 (仅作初值，与 C# NaN→中位数 语义一致)
Private Sub SvBuildRequestRow(ByRef data As Variant, ByRef schema As TSvSchema, _
                              ByVal arrayRow As Long, ByVal useReq As Boolean, _
                              ByRef request As Variant, ByVal reqRow As Long, _
                              ByRef map() As Long, ByRef med() As Double, _
                              ByRef feature() As Double, ByVal label As String)
    Dim rLo As Long, cLo As Long, rrLo As Long, rcLo As Long
    Dim c As Long, dataCol As Long, mapIdx As Long
    Dim cell As Variant, d As Double
    Dim isInc As Boolean, isFix As Boolean

    rLo = LBound(data, 1): cLo = LBound(data, 2)
    If useReq Then
        rrLo = LBound(request, 1): rcLo = LBound(request, 2)
    End If
    For c = 0 To schema.FeatureCount - 1
        dataCol = schema.FeatureCols(c)
        If useReq Then
            mapIdx = map(dataCol)
            If mapIdx >= 0 Then
                cell = request(rrLo + reqRow, rcLo + mapIdx)
            Else
                cell = Empty
            End If
        Else
            cell = data(rLo + arrayRow, cLo + dataCol)
        End If
        isInc = SvIsIncoming(schema, dataCol)
        isFix = SvIsFixed(schema, dataCol)
        If SvIsBlank(cell) Then
            If isInc Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "Incoming column '" & schema.Headers(dataCol) & "' is blank in " & label & "."
            End If
            feature(c) = med(c)
        ElseIf Not SvIsNumericCell(cell) Then
            If isInc Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "Incoming column '" & schema.Headers(dataCol) & "' is non-numeric in " & label & "."
            ElseIf isFix Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "Fixed column '" & schema.Headers(dataCol) & "' is non-numeric in " & label & "."
            Else
                feature(c) = med(c)
            End If
        Else
            d = SvToDoubleCell(cell)
            feature(c) = d
        End If
    Next c
End Sub

Private Sub SvMapRequestColumns(ByRef dataSchema As TSvSchema, ByRef reqSchema As TSvSchema, _
                                ByRef map() As Long)
    Dim dc As Long, rc As Long, found As Long
    For dc = 0 To 127: map(dc) = -1: Next dc
    For rc = 0 To reqSchema.ColCount - 1
        If SvClassifyHeader(reqSchema.Headers(rc)) <> "" Then
            found = -1
            For dc = 0 To dataSchema.ColCount - 1
                If StrComp(dataSchema.Headers(dc), reqSchema.Headers(rc), vbTextCompare) = 0 Then
                    found = dc: Exit For
                End If
            Next dc
            If found < 0 Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "Request table column '" & reqSchema.Headers(rc) & "' does not match any data column."
            End If
            map(found) = rc
        End If
    Next rc
End Sub

' 边界: 默认历史 min/max; bounds 支持 2 列 (按可调顺序) 或 3 列 (名称/序号, 下界, 上界)
Private Sub SvBuildBounds(ByRef schema As TSvSchema, ByRef X() As Double, ByVal n As Long, _
                          ByRef bounds As Variant, ByRef bnd() As Double)
    Dim v As Long: v = schema.VariableCount
    Dim c As Long, i As Long, p As Long, r As Long
    Dim bR0 As Long, bC0 As Long, bRows As Long, bCols As Long
    Dim key As Variant, name As String, idx As Long
    Dim lo As Double, hi As Double

    ReDim bnd(0 To v - 1, 0 To 1)
    For c = 0 To v - 1
        p = SvFeaturePos(schema, schema.Variable(c))
        bnd(c, 0) = X(0, p): bnd(c, 1) = X(0, p)
        For i = 1 To n - 1
            If X(i, p) < bnd(c, 0) Then bnd(c, 0) = X(i, p)
            If X(i, p) > bnd(c, 1) Then bnd(c, 1) = X(i, p)
        Next i
    Next c

    If IsEmpty(bounds) Or IsNull(bounds) Then Exit Sub
    If Not IsArray(bounds) Then Exit Sub
    bR0 = LBound(bounds, 1): bC0 = LBound(bounds, 2)
    bRows = UBound(bounds, 1) - bR0 + 1
    bCols = UBound(bounds, 2) - bC0 + 1
    If bRows = 1 And bCols = 1 Then
        If SvIsBlank(bounds(bR0, bC0)) Then Exit Sub
    End If
    If bRows = 0 Then Exit Sub
    If bCols = 2 Then
        If bRows > v Then
            Err.Raise ERR_SV_BAD_BOUNDS, "SolveUtils", _
                "Bounds table has " & bRows & " rows but there are only " & v & " variable columns."
        End If
        For r = 0 To bRows - 1
            lo = SvToDoubleCell(bounds(bR0 + r, bC0))
            hi = SvToDoubleCell(bounds(bR0 + r, bC0 + 1))
            SvValidateBoundPair lo, hi, r + 1
            bnd(r, 0) = lo: bnd(r, 1) = hi
        Next r
        Exit Sub
    End If
    If bCols = 3 Then
        For r = 0 To bRows - 1
            key = bounds(bR0 + r, bC0)
            If SvIsBlank(key) Then
                Err.Raise ERR_SV_BAD_BOUNDS, "SolveUtils", _
                    "Bounds table row " & (r + 1) & " is missing the variable name/index."
            End If
            If SvIsNumericCell(key) Then
                idx = CLng(CDbl(key))
                If idx < 1 Or idx > v Then
                    Err.Raise ERR_SV_BAD_BOUNDS, "SolveUtils", _
                        "Bounds table row " & (r + 1) & ": variable index " & idx & " is out of range [1," & v & "]."
                End If
                idx = idx - 1
            Else
                name = Trim$(CStr(key))
                idx = -1
                For c = 0 To v - 1
                    If StrComp(schema.Headers(schema.Variable(c)), name, vbTextCompare) = 0 Then
                        idx = c: Exit For
                    End If
                Next c
                If idx < 0 Then
                    Err.Raise ERR_SV_BAD_BOUNDS, "SolveUtils", _
                        "Bounds table row " & (r + 1) & ": variable '" & name & "' does not match any variable column."
                End If
            End If
            lo = SvToDoubleCell(bounds(bR0 + r, bC0 + 1))
            hi = SvToDoubleCell(bounds(bR0 + r, bC0 + 2))
            SvValidateBoundPair lo, hi, r + 1
            bnd(idx, 0) = lo: bnd(idx, 1) = hi
        Next r
        Exit Sub
    End If
    Err.Raise ERR_SV_BAD_BOUNDS, "SolveUtils", _
        "Bounds table must have 2 (by variable order) or 3 (name/index, lower, upper) columns; got " & bCols & "."
End Sub

Private Sub SvValidateBoundPair(ByVal lo As Double, ByVal hi As Double, ByVal row As Long)
    If Not SvIsFinite(lo) Or Not SvIsFinite(hi) Then
        Err.Raise ERR_SV_BAD_BOUNDS, "SolveUtils", "Bounds table row " & row & " contains a non-finite bound."
    End If
    If lo > hi Then
        Err.Raise ERR_SV_BAD_BOUNDS, "SolveUtils", _
            "Bounds table row " & row & " has lower > upper (" & CStr(lo) & " > " & CStr(hi) & ")."
    End If
End Sub

' auto → 每输出 CV 择优; 显式 → linear/poly
Private Sub SvFitModels(ByRef X() As Double, ByVal n As Long, ByVal k As Long, _
                        ByRef Y() As Double, ByVal outCount As Long, _
                        ByVal kind As String, ByVal seed As Long, _
                        ByRef models() As TSvModel)
    Dim j As Long, i As Long
    Dim yj() As Double
    Dim chosen As String, scheme As String, r2 As Double, mae As Double, polySkip As Boolean

    ReDim models(0 To outCount - 1)
    For j = 0 To outCount - 1
        ReDim yj(0 To n - 1)
        For i = 0 To n - 1: yj(i) = Y(i, j): Next i
        If kind = MODEL_AUTO Then
            SvFitAuto X, n, k, yj, seed, models(j), chosen, scheme, r2, mae, polySkip
        Else
            SvFitExpanded X, n, k, yj, kind, (kind = MODEL_POLY), models(j)
        End If
    Next j
End Sub

' UDF 契约表: [请求行, 可调列名…, 输出预测…, 最大偏差σ, 状态] (v1 无速率块)
Public Function SolveInverse(ByRef data As Variant, ByRef request As Variant, ByRef bounds As Variant, _
                           ByVal model As Variant, ByVal seed As Long, ByVal maxStarts As Long) As Variant
    Dim schema As TSvSchema, reqSchema As TSvSchema
    Dim mdl As String
    Dim n As Long, k As Long, outCount As Long, v As Long
    Dim X() As Double, Y() As Double
    Dim med() As Double
    Dim models() As TSvModel
    Dim varPos() As Long
    Dim bnd() As Double
    Dim requests() As Double, targets() As Double, tgtMask() As Boolean
    Dim labels() As String
    Dim r As Long, q As Long, i As Long, c As Long, j As Long
    Dim useReq As Boolean
    Dim core() As Double
    Dim tbl() As Variant
    Dim width As Long
    Dim map(0 To 127) As Long
    Dim rLo As Long, cLo As Long, rrLo As Long, rcLo As Long
    Dim rows As Long
    Dim cell As Variant
    Dim anyTgt As Boolean
    Dim feature() As Double
    Dim arrayRow As Long, reqRow As Long, emitted As Long
    Dim label As String
    Dim d As Double

    SvCoerceOptionalTableArg request, "request"
    SvCoerceOptionalTableArg bounds, "bounds"
    SvParseSchema data, schema
    If schema.HistoryCount = 0 Then
        Err.Raise ERR_SV_NO_HISTORY, "SolveUtils", _
            "Data table contains no history rows (every row has a blank adjustable column)."
    End If
    mdl = SvNormalizeModel(model)
    n = schema.HistoryCount: k = schema.FeatureCount
    outCount = schema.OutputCount: v = schema.VariableCount
    If maxStarts < MIN_STARTS Or maxStarts > MAX_STARTS_LIMIT Then
        Err.Raise ERR_SV_LIMIT, "SolveUtils", _
            "max_starts must be between " & MIN_STARTS & " and " & MAX_STARTS_LIMIT & " (got " & maxStarts & ")."
    End If
    SvBuildHistory data, schema, X, Y
    SvFeatureMedians X, n, k, med

    useReq = Not SvIsBlank(request)
    If useReq Then
        If schema.RequestCount > 0 Then
            Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                "Data table must not contain request rows when a separate request table is provided."
        End If
        SvParseSchema request, reqSchema, False
        SvMapRequestColumns schema, reqSchema, map
        ' 统计非空请求行
        rrLo = LBound(request, 1)
        rows = UBound(request, 1) - rrLo + 1
        r = 0
        For i = 1 To rows - 1
            anyTgt = False
            For c = 0 To UBound(request, 2) - LBound(request, 2)
                If Not SvIsBlank(request(rrLo + i, LBound(request, 2) + c)) Then anyTgt = True: Exit For
            Next c
            If anyTgt Then r = r + 1
        Next i
    Else
        r = schema.RequestCount
    End If
    If r = 0 Then
        Err.Raise ERR_SV_NO_REQUEST, "SolveUtils", _
            "No request rows found. Leave at least one adjustable column blank (or pass a request table)."
    End If
    If r > MAX_REQUEST_ROWS Then
        Err.Raise ERR_SV_LIMIT, "SolveUtils", "Too many request rows: " & r & " (limit " & MAX_REQUEST_ROWS & ")."
    End If

    ReDim requests(0 To r - 1, 0 To k - 1)
    ReDim targets(0 To r - 1, 0 To outCount - 1)
    ReDim tgtMask(0 To r - 1, 0 To outCount - 1)
    ReDim labels(0 To r - 1)
    ReDim feature(0 To k - 1)

    rLo = LBound(data, 1): cLo = LBound(data, 2)
    If Not useReq Then
        For q = 0 To r - 1
            arrayRow = schema.RequestRows(q) + 1
            label = "data row " & (schema.RequestRows(q) + 2)
            SvBuildRequestRow data, schema, arrayRow, False, Empty, 0, map, med, feature, label
            For c = 0 To k - 1: requests(q, c) = feature(c): Next c
            anyTgt = False
            For j = 0 To outCount - 1
                cell = data(rLo + arrayRow, cLo + schema.Output(j))
                If SvIsBlank(cell) Then
                    tgtMask(q, j) = False
                Else
                    d = SvToDoubleCell(cell)
                    targets(q, j) = d
                    tgtMask(q, j) = True
                    anyTgt = True
                End If
            Next j
            If Not anyTgt Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "Request row (data row " & (schema.RequestRows(q) + 2) & ") has no output target."
            End If
            labels(q) = "数据第" & (schema.RequestRows(q) + 2) & "行"
        Next q
    Else
        rrLo = LBound(request, 1): rcLo = LBound(request, 2)
        emitted = 0
        For i = 1 To rows - 1
            anyTgt = False
            For c = 0 To UBound(request, 2) - LBound(request, 2)
                If Not SvIsBlank(request(rrLo + i, rcLo + c)) Then anyTgt = True: Exit For
            Next c
            If Not anyTgt Then GoTo NextReqRow
            reqRow = i
            label = "request table row " & (i + 1)
            SvBuildRequestRow data, schema, 0, True, request, reqRow, map, med, feature, label
            For c = 0 To k - 1: requests(emitted, c) = feature(c): Next c
            anyTgt = False
            For j = 0 To outCount - 1
                If map(schema.Output(j)) >= 0 Then
                    cell = request(rrLo + reqRow, rcLo + map(schema.Output(j)))
                Else
                    cell = Empty
                End If
                If SvIsBlank(cell) Then
                    tgtMask(emitted, j) = False
                Else
                    d = SvToDoubleCell(cell)
                    targets(emitted, j) = d
                    tgtMask(emitted, j) = True
                    anyTgt = True
                End If
            Next j
            If Not anyTgt Then
                Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                    "Request table row " & (i + 1) & " has no output target."
            End If
            emitted = emitted + 1
            labels(emitted - 1) = "请求" & emitted
NextReqRow:
        Next i
    End If

    SvFitModels X, n, k, Y, outCount, mdl, seed, models
    ReDim varPos(0 To v - 1)
    For c = 0 To v - 1: varPos(c) = SvFeaturePos(schema, schema.Variable(c)): Next c
    SvBuildBounds schema, X, n, bounds, bnd
    SvSolveInverseCore X, n, Y, outCount, varPos, requests, r, targets, tgtMask, bnd, models, seed, maxStarts, core

    width = 1 + v + outCount + 2
    ReDim tbl(1 To r + 1, 1 To width)
    tbl(1, 1) = "请求行"
    For c = 0 To v - 1: tbl(1, 2 + c) = schema.Headers(schema.Variable(c)): Next c
    For j = 0 To outCount - 1: tbl(1, 2 + v + j) = schema.Headers(schema.Output(j)) & "预测": Next j
    tbl(1, 2 + v + outCount) = "最大偏差σ"
    tbl(1, width) = "状态"
    For q = 0 To r - 1
        tbl(2 + q, 1) = labels(q)
        For c = 0 To v - 1: tbl(2 + q, 2 + c) = core(q, c): Next c
        For j = 0 To outCount - 1: tbl(2 + q, 2 + v + j) = core(q, v + j): Next j
        tbl(2 + q, 2 + v + outCount) = core(q, v + outCount)
        If core(q, v + outCount + 1) = 0# Then
            tbl(2 + q, width) = STATUS_REACHABLE
        Else
            tbl(2 + q, width) = STATUS_UNREACHABLE
        End If
    Next q
    SolveInverse = tbl
End Function

' C# G6 (6 位有效数字, 区域无关) 的 VBA 近似实现
Private Function SvNumG6(ByVal v As Double) As String
    Dim av As Double, e As Long, dec As Long
    Dim s As String, mant As String, ex As String
    Dim ep As Long

    If v = 0# Then
        SvNumG6 = "0"
        Exit Function
    End If
    av = Abs(v)
    If av >= 0.0001 And av < 1000000# Then
        e = Int(Log(av) / Log(10#))
        If e < -4 Then e = -4
        If e > 5 Then e = 5
        dec = 5 - e
        s = Format$(v, "0." & String$(dec, "0"))
        If InStr(s, ".") > 0 Then
            Do While Right$(s, 1) = "0"
                s = Left$(s, Len(s) - 1)
            Loop
            If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
        End If
    Else
        s = Format$(v, "0.#####E+00")
        ep = InStr(s, "E")
        If ep > 0 Then
            mant = Left$(s, ep - 1): ex = Mid$(s, ep)
            If InStr(mant, ".") > 0 Then
                Do While Right$(mant, 1) = "0"
                    mant = Left$(mant, Len(mant) - 1)
                Loop
                If Right$(mant, 1) = "." Then mant = Left$(mant, Len(mant) - 1)
            End If
            s = mant & ex
        End If
    End If
    SvNumG6 = Replace(s, ",", ".")
End Function

Private Function SvTermPowerSum(ByRef model As TSvModel, ByVal term As Long) As Long
    Dim j As Long, s As Long
    For j = 1 To model.BaseCount: s = s + model.Powers(term, j): Next j
    SvTermPowerSum = s
End Function

Private Function SvTermFirstBase(ByRef model As TSvModel, ByVal term As Long) As Long
    Dim j As Long
    For j = 1 To model.BaseCount
        If model.Powers(term, j) = 1 Then SvTermFirstBase = j - 1: Exit Function
    Next j
    SvTermFirstBase = -1
End Function

Private Function SvTermFactor(ByRef model As TSvModel, ByVal term As Long, ByRef names() As String) As String
    Dim parts(0 To 31) As String, cnt As Long
    Dim j As Long, p As Long
    For j = 1 To model.BaseCount
        p = model.Powers(term, j)
        If p > 0 Then
            If p = 1 Then
                parts(cnt) = names(j - 1)
            Else
                parts(cnt) = names(j - 1) & "^" & CStr(p)
            End If
            cnt = cnt + 1
        End If
    Next j
    If cnt = 0 Then
        SvTermFactor = ""
    ElseIf cnt = 1 Then
        SvTermFactor = parts(0)
    Else
        Dim t() As String
        ReDim t(0 To cnt - 1)
        For j = 0 To cnt - 1: t(j) = parts(j): Next j
        SvTermFactor = Join(t, "*")
    End If
End Function

Private Function SvExpression(ByRef model As TSvModel, ByRef names() As String) As String
    ' 每项写 2 槽 (符号 + 因子), 容量由 POLY_TERM_LIMIT 派生 (≥65 项时旧缓冲 128 越界)
    Dim parts(0 To 2 * POLY_TERM_LIMIT) As String, cnt As Long
    Dim t As Long, absC As Double, factor As String
    Dim c As Double
    Dim used() As String

    For t = 1 To model.TermCount
        c = model.Coef(t)
        If c <> 0# Then
            If c > 0# Then parts(cnt) = " + " Else parts(cnt) = " - "
            cnt = cnt + 1
            absC = Abs(c)
            factor = SvTermFactor(model, t, names)
            If Len(factor) = 0 Then
                parts(cnt) = SvNumG6(absC)
            ElseIf absC = 1# Then
                parts(cnt) = factor
            Else
                parts(cnt) = SvNumG6(absC) & "*" & factor
            End If
            cnt = cnt + 1
        End If
    Next t
    If cnt > 0 Then
        ReDim used(0 To cnt - 1)
        For t = 0 To cnt - 1: used(t) = parts(t): Next t
        SvExpression = SvNumG6(model.Intercept) & Join(used, "")
    Else
        SvExpression = SvNumG6(model.Intercept)
    End If
End Function

Private Function SvInverseFormula(ByRef model As TSvModel, ByRef names() As String, _
                                  ByVal varIdx As Long, ByVal outputName As String) As String
    Dim t As Long, j As Long, coefVar As Double, c As Double, absC As Double
    Dim s As String

    coefVar = 0#
    For t = 1 To model.TermCount
        If model.Powers(t, varIdx + 1) = 1 And SvTermPowerSum(model, t) = 1 Then
            coefVar = model.Coef(t): Exit For
        End If
    Next t
    If coefVar = 0# Then
        SvInverseFormula = ""
        Exit Function
    End If
    s = names(varIdx) & " = (" & outputName
    If model.Intercept > 0# Then
        s = s & " - " & SvNumG6(model.Intercept)
    ElseIf model.Intercept < 0# Then
        s = s & " + " & SvNumG6(-model.Intercept)
    End If
    For t = 1 To model.TermCount
        If SvTermPowerSum(model, t) = 1 And model.Powers(t, varIdx + 1) <> 1 Then
            j = SvTermFirstBase(model, t)
            If j >= 0 Then
                c = model.Coef(t)
                If c <> 0# Then
                    If c > 0# Then s = s & " - " Else s = s & " + "
                    absC = Abs(c)
                    If absC = 1# Then
                        s = s & names(j)
                    Else
                        s = s & SvNumG6(absC) & "*" & names(j)
                    End If
                End If
            End If
        End If
    Next t
    s = s & ") / " & SvNumG6(coefVar)
    SvInverseFormula = s
End Function

' 质量表: [输出, 候选, CV方案, CV_R2, CV_MAE, 选用] (v1: linear/poly)
Public Function SolveQuality(ByRef data As Variant, ByVal model As Variant, ByVal seed As Long) As Variant
    Dim schema As TSvSchema
    Dim mdl As String
    Dim X() As Double, Y() As Double
    Dim n As Long, k As Long, outCount As Long
    Dim i As Long, j As Long, c As Long
    Dim yj() As Double
    Dim rows() As Variant
    Dim rowCount As Long
    Dim linScheme As String, linR2 As Double, linMae As Double, linSkipped As Boolean
    Dim polyScheme As String, polyR2 As Double, polyMae As Double, polySkipped As Boolean
    Dim firstErr As Long, firstDesc As String
    Dim best As String, bestR2 As Double
    Dim outName As String
    Dim scheme As String, r2 As Double, mae As Double
    Dim tbl() As Variant

    SvParseSchema data, schema
    If schema.HistoryCount = 0 Then
        Err.Raise ERR_SV_NO_HISTORY, "SolveUtils", "Data table contains no history rows."
    End If
    mdl = SvNormalizeModel(model)
    n = schema.HistoryCount: k = schema.FeatureCount: outCount = schema.OutputCount
    SvBuildHistory data, schema, X, Y

    ReDim rows(0 To outCount * 2 - 1, 0 To 5)
    rowCount = 0
    For j = 0 To outCount - 1
        ReDim yj(0 To n - 1)
        For i = 0 To n - 1: yj(i) = Y(i, j): Next i
        outName = schema.Headers(schema.Output(j))

        If mdl = MODEL_AUTO Then
            firstErr = 0: firstDesc = ""
            linSkipped = False: polySkipped = False
            On Error Resume Next
            Err.Clear
            SvCrossValidate X, n, k, yj, MODEL_LINEAR, seed, linScheme, linR2, linMae
            If Err.Number <> 0 Then
                linSkipped = True
                firstErr = Err.Number: firstDesc = Err.Description
                Err.Clear
            End If
            On Error GoTo 0

            polySkipped = SvPolyExcludedBySample(n, k)
            If Not polySkipped Then
                On Error Resume Next
                Err.Clear
                SvCrossValidate X, n, k, yj, MODEL_POLY, seed, polyScheme, polyR2, polyMae
                If Err.Number <> 0 Then
                    polySkipped = True
                    If firstErr = 0 Then firstErr = Err.Number: firstDesc = Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            End If
            If linSkipped And polySkipped Then
                If firstErr <> 0 Then
                    Err.Raise firstErr, "SolveUtils", firstDesc
                Else
                    Err.Raise ERR_SV_CV, "SolveUtils", "No model candidate could be cross-validated."
                End If
            End If

            If linSkipped Then
                best = MODEL_POLY: bestR2 = polyR2
            Else
                best = MODEL_LINEAR: bestR2 = linR2
            End If
            If Not polySkipped Then
                If polyR2 > bestR2 + R2_TIE_TOL Then
                    best = MODEL_POLY: bestR2 = polyR2
                End If
            End If

            If linSkipped Then
                rows(rowCount, 0) = outName: rows(rowCount, 1) = MODEL_LINEAR
                rows(rowCount, 2) = SCHEME_SKIP: rows(rowCount, 3) = Empty
                rows(rowCount, 4) = Empty: rows(rowCount, 5) = "否"
            Else
                rows(rowCount, 0) = outName: rows(rowCount, 1) = MODEL_LINEAR
                rows(rowCount, 2) = linScheme: rows(rowCount, 3) = linR2
                rows(rowCount, 4) = linMae
                If best = MODEL_LINEAR Then rows(rowCount, 5) = "是" Else rows(rowCount, 5) = "否"
            End If
            rowCount = rowCount + 1

            If polySkipped Then
                rows(rowCount, 0) = outName: rows(rowCount, 1) = MODEL_POLY
                rows(rowCount, 2) = SCHEME_SKIP: rows(rowCount, 3) = Empty
                rows(rowCount, 4) = Empty: rows(rowCount, 5) = "否"
            Else
                rows(rowCount, 0) = outName: rows(rowCount, 1) = MODEL_POLY
                rows(rowCount, 2) = polyScheme: rows(rowCount, 3) = polyR2
                rows(rowCount, 4) = polyMae
                If best = MODEL_POLY Then rows(rowCount, 5) = "是" Else rows(rowCount, 5) = "否"
            End If
            rowCount = rowCount + 1
        Else
            SvCrossValidate X, n, k, yj, mdl, seed, scheme, r2, mae
            rows(rowCount, 0) = outName: rows(rowCount, 1) = mdl
            rows(rowCount, 2) = scheme: rows(rowCount, 3) = r2
            rows(rowCount, 4) = mae: rows(rowCount, 5) = "是"
            rowCount = rowCount + 1
        End If
    Next j

    ReDim tbl(1 To rowCount + 1, 1 To 6)
    tbl(1, 1) = "输出": tbl(1, 2) = "候选": tbl(1, 3) = "CV方案"
    tbl(1, 4) = "CV_R2": tbl(1, 5) = "CV_MAE": tbl(1, 6) = "选用"
    For i = 0 To rowCount - 1
        For c = 0 To 5: tbl(2 + i, 1 + c) = rows(i, c): Next c
    Next i
    SolveQuality = tbl
End Function

' 前向预测表: values (n×k, 列序 = 数据特征序) → N×M 数值矩阵 (1-based)
Public Function SolvePredict(ByRef data As Variant, ByRef values As Variant, ByVal model As Variant) As Variant
    Dim schema As TSvSchema
    Dim mdl As String
    Dim X() As Double, Y() As Double
    Dim n As Long, k As Long, outCount As Long
    Dim i As Long, c As Long, j As Long
    Dim med() As Double
    Dim models() As TSvModel
    Dim vR0 As Long, vC0 As Long, vRows As Long, vCols As Long
    Dim rLo As Long, cLo As Long
    Dim cell As Variant
    Dim feature() As Double
    Dim tbl() As Variant
    Dim dataCol As Long

    SvParseSchema data, schema
    If schema.HistoryCount = 0 Then
        Err.Raise ERR_SV_NO_HISTORY, "SolveUtils", "Data table contains no history rows."
    End If
    mdl = SvNormalizeModel(model)
    n = schema.HistoryCount: k = schema.FeatureCount: outCount = schema.OutputCount
    SvBuildHistory data, schema, X, Y
    SvFeatureMedians X, n, k, med

    SvCoerceTableArg values, "values"
    If Not IsArray(values) Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "Values table must be a 2D range or array."
    End If
    vR0 = LBound(values, 1): vC0 = LBound(values, 2)
    vRows = UBound(values, 1) - vR0 + 1
    vCols = UBound(values, 2) - vC0 + 1
    If vCols <> k Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
            "Values table must have " & k & " columns (Incoming + Variable + Fixed, in data order); got " & vCols & "."
    End If
    If vRows < 1 Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "Values table must contain at least one row."
    End If
    ' C# PredictTable 对 auto 固定使用 seed=42
    SvFitModels X, n, k, Y, outCount, mdl, 42, models

    rLo = LBound(data, 1): cLo = LBound(data, 2)
    ReDim feature(0 To k - 1)
    ReDim tbl(1 To vRows, 1 To outCount)
    For i = 0 To vRows - 1
        For c = 0 To k - 1
            dataCol = schema.FeatureCols(c)
            cell = values(vR0 + i, vC0 + c)
            If SvIsBlank(cell) Then
                If SvIsFixed(schema, dataCol) Then
                    feature(c) = med(c)
                Else
                    Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", _
                        "Values row " & (i + 1) & " is missing a value in column '" & schema.Headers(dataCol) & "'."
                End If
            Else
                feature(c) = SvToDoubleCell(cell)
            End If
        Next c
        For j = 0 To outCount - 1
            tbl(1 + i, 1 + j) = SvPredict(models(j), feature)
        Next j
    Next i
    SolvePredict = tbl
End Function

' 方程表: [输出, 类型, 表达式]; 单可调线性模型额外输出反解公式
Public Function SolveEquation(ByRef data As Variant, ByVal model As Variant) As Variant
    Dim schema As TSvSchema
    Dim mdl As String
    Dim X() As Double, Y() As Double
    Dim n As Long, k As Long, outCount As Long, v As Long
    Dim i As Long, j As Long, c As Long
    Dim yj() As Double
    Dim models() As TSvModel
    Dim names() As String
    Dim rows() As Variant
    Dim rowCount As Long
    Dim outName As String, varName As String
    Dim varIdx As Long
    Dim inv As String
    Dim tbl() As Variant

    SvParseSchema data, schema
    If schema.HistoryCount = 0 Then
        Err.Raise ERR_SV_NO_HISTORY, "SolveUtils", "Data table contains no history rows."
    End If
    mdl = SvNormalizeModel(model)
    n = schema.HistoryCount: k = schema.FeatureCount
    outCount = schema.OutputCount: v = schema.VariableCount
    SvBuildHistory data, schema, X, Y
    ' C# Equation 对 auto 固定使用 seed=42
    SvFitModels X, n, k, Y, outCount, mdl, 42, models

    ReDim names(0 To k - 1)
    For c = 0 To k - 1: names(c) = schema.Headers(schema.FeatureCols(c)): Next c
    ReDim rows(0 To outCount * 3 - 1, 0 To 2)
    rowCount = 0
    For j = 0 To outCount - 1
        outName = schema.Headers(schema.Output(j))
        rows(rowCount, 0) = outName
        rows(rowCount, 1) = "前向方程"
        rows(rowCount, 2) = outName & " = " & SvExpression(models(j), names)
        rowCount = rowCount + 1
        If models(j).Kind = MODEL_LINEAR And v = 1 Then
            varIdx = SvFeaturePos(schema, schema.Variable(0))
            varName = names(varIdx)
            inv = SvInverseFormula(models(j), names, varIdx, outName)
            If Len(inv) > 0 Then
                rows(rowCount, 0) = outName
                rows(rowCount, 1) = "反解公式"
                rows(rowCount, 2) = inv
                rowCount = rowCount + 1
            End If
        End If
        If models(j).RankDeficient Then
            rows(rowCount, 0) = outName
            rows(rowCount, 1) = "备注"
            rows(rowCount, 2) = "特征共线: 部分系数不可辨识 (已置 0); 建议使用岭回归或移除冗余特征。"
            rowCount = rowCount + 1
        End If
    Next j

    ReDim tbl(1 To rowCount + 1, 1 To 3)
    tbl(1, 1) = "输出": tbl(1, 2) = "类型": tbl(1, 3) = "表达式"
    For i = 0 To rowCount - 1
        For c = 0 To 2: tbl(2 + i, 1 + c) = rows(i, c): Next c
    Next i
    SolveEquation = tbl
End Function

'=============================================================================
' Phase 2 — UDF 层 (参数全 Variant; 失败返回 CVErr(xlErrValue); Range/数组双路径)
'=============================================================================

Private Function SvNormTable(ByVal v As Variant, ByVal name As String) As Variant
    Dim n As Variant
    If IsMissing(v) Or IsEmpty(v) Or IsNull(v) Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "'" & name & "' must be a range or array, not empty."
    End If
    SvEnsureCore
    n = VK.Normalize2D(v)
    If IsError(n) Or Not IsArray(n) Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "'" & name & "' must be a 2D range or array."
    End If
    SvNormTable = n
End Function

Private Function SvNormOpt(ByVal v As Variant) As Variant
    Dim n As Variant
    If IsMissing(v) Or IsEmpty(v) Or IsNull(v) Then
        SvNormOpt = Empty
        Exit Function
    End If
    If Not IsArray(v) And Not IsObject(v) Then
        SvNormOpt = Empty
        Exit Function
    End If
    SvEnsureCore
    n = VK.Normalize2D(v)
    If IsError(n) Or Not IsArray(n) Then
        SvNormOpt = Empty
    Else
        SvNormOpt = n
    End If
End Function

Private Function SvOptLong(ByVal v As Variant, ByVal dflt As Long) As Long
    If IsMissing(v) Or IsEmpty(v) Or IsNull(v) Then
        SvOptLong = dflt
    Else
        SvOptLong = CLng(CDbl(v))
    End If
End Function

Private Function SvOptModel(ByVal v As Variant) As Variant
    If IsMissing(v) Or IsEmpty(v) Or IsNull(v) Then
        SvOptModel = Empty
    Else
        SvOptModel = CStr(v)
    End If
End Function

' 双路径: Range 对象 → 1-based 2D Variant 数组; 数组原样 (任意 LBound)
Private Sub SvCoerceTableArg(ByRef v As Variant, ByVal name As String)
    Dim n As Variant
    If Not IsObject(v) Then Exit Sub
    If Not TypeOf v Is Range Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "'" & name & "' must be a 2D range or array."
    End If
    SvEnsureCore
    n = VK.Normalize2D(v)
    If IsError(n) Or Not IsArray(n) Then
        Err.Raise ERR_SV_INVALID_INPUT, "SolveUtils", "'" & name & "' must be a 2D range or array."
    End If
    v = n
End Sub

Private Sub SvCoerceOptionalTableArg(ByRef v As Variant, ByVal name As String)
    If IsObject(v) Then SvCoerceTableArg v, name
End Sub

Public Function UDF_SOLVE_INVERSE(ByVal data As Variant, Optional ByVal request As Variant, _
                                  Optional ByVal bounds As Variant, Optional ByVal model As Variant, _
                                  Optional ByVal seed As Variant, _
                                  Optional ByVal max_starts As Variant) As Variant
    On Error GoTo EH
    UDF_SOLVE_INVERSE = SolveInverse(SvNormTable(data, "data"), SvNormOpt(request), SvNormOpt(bounds), _
                                  SvOptModel(model), SvOptLong(seed, 42), SvOptLong(max_starts, 10))
    Exit Function
EH:
    UDF_SOLVE_INVERSE = CVErr(xlErrValue)
End Function

Public Function UDF_SOLVE_PREDICT(ByVal data As Variant, ByVal values As Variant, _
                                  Optional ByVal model As Variant) As Variant
    On Error GoTo EH
    UDF_SOLVE_PREDICT = SolvePredict(SvNormTable(data, "data"), SvNormTable(values, "values"), _
                                       SvOptModel(model))
    Exit Function
EH:
    UDF_SOLVE_PREDICT = CVErr(xlErrValue)
End Function

Public Function UDF_SOLVE_QUALITY(ByVal data As Variant, Optional ByVal model As Variant, _
                                  Optional ByVal seed As Variant) As Variant
    On Error GoTo EH
    UDF_SOLVE_QUALITY = SolveQuality(SvNormTable(data, "data"), SvOptModel(model), SvOptLong(seed, 42))
    Exit Function
EH:
    UDF_SOLVE_QUALITY = CVErr(xlErrValue)
End Function

Public Function UDF_SOLVE_EQUATION(ByVal data As Variant, Optional ByVal model As Variant) As Variant
    On Error GoTo EH
    UDF_SOLVE_EQUATION = SolveEquation(SvNormTable(data, "data"), SvOptModel(model))
    Exit Function
EH:
    UDF_SOLVE_EQUATION = CVErr(xlErrValue)
End Function
