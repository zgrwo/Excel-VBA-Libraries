Option Explicit

'==============================================================================
' Module:       StatsUtils
' Purpose:      Statistics: descriptive, inference, distribution, correlation
' Layer:        Statistics
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       77 functions/subs
'==============================================================================

Private DP As New DictProxy

'=====================================================================
' StatsUtils.bas — 统计函数 (集中趋势/离散度/形态/排名/关联/变换/推断/分布)
'
' 工作表函数 (UDF_STAT_*):
'   UDF_STAT_MEAN/MEDIAN/MODE/MINMAX — 集中趋势
'   UDF_STAT_GEOMEAN/HARMEAN/TRIMEAN/RMS — 其他均值
'   UDF_STAT_STDEV/STDEVP/VAR/VARP — 离散度
'   UDF_STAT_PERCENTILE/IQR/SKEW/KURTOSIS — 分布
'   UDF_STAT_COV/CORREL/R2 — 关联
'   UDF_STAT_ZSCORE/NORMALIZE/WINSORIZE/MA/LININTERP — 变换
'   UDF_STAT_RANK/PERCENTRANK/SE/MAD — 排名/误差
'
' VBA-only (返回 Dictionary/Object + Err.Raise):
'   ConfidenceInterval, Binning
'
' 错误约定:
'   无效输入 (数据空/长度不等)  →  CVErr(xlErrValue) / #VALUE!
'   数学未定义 (零方差等)       →  CVErr(xlErrDiv0)  / #DIV/0!
'   无效参数 (负数对数等)       →  CVErr(xlErrNum)   / #NUM!
'   无结果 (无众数等)           →  CVErr(xlErrNA)    / #N/A
'
' 基础描述:
'   Mean           — 算术均值
'   GeometricMean  — 几何均值
'   HarmonicMean   — 调和均值
'   TrimMean       — 截尾均值
'   RootMeanSquare — 均方根
'   MeanAbsDev     — 平均绝对偏差
'   Median         — 中位数
'   Mode           — 众数 (唯一值时返回 #N/A)
'   MinMax         — 同时返回最小值/最大值
'
' 离散度:
'   StdDev / StdDevP       — 标准差 (样本/总体)
'   Variance / VarianceP   — 方差 (样本/总体)
'   Percentile             — 百分位数
'   IQR                    — 四分位距
'   Skewness               — 偏度
'   Kurtosis               — 超额峰度
'
' 推断/诊断:
'   StandardError          — 标准误差
'   ConfidenceInterval     — 置信区间 (VBA-only, 返回 Dictionary)
'   Covariance             — 协方差
'   Correlation            — 相关系数
'   RSquare                — R² 判定系数
'
' 排名:
'   Rank                   — 值在数据集中的排名
'   PercentRank            — 百分位排名
'
' 数据处理:
'   ZScore                 — Z 标准化
'   Normalize              — Min-Max 归一化
'   Winsorize              — 缩尾处理
'   MovingAverage          — 移动平均
'   Binning                — 等宽分箱 (VBA-only, 返回 Dictionary)
'   LinInterp              — 线性插值
'
' 内部辅助:
'   ExtractDoubles       — 从 Range/1D 数组中提取 Double 数组
'   ExtractPairedDoubles — 成对提取两列, 仅保留双列均为数值的行
'   MeanDouble           — 对预提取的 Double() 数组计算均值
'   StdDevDouble         — 对预提取的 Double() 数组计算样本标准差
'   CovarianceDouble     — 对预提取的 Double() 数组计算样本协方差
'   QuickSortDouble      — Double 数组快速排序 (含插入排序优化)
'   InsertionSortDouble  — 小数组插入排序
'   QuickPercentile      — 对已排序数组计算百分位数
'   TDistCritical        — t 分布临界值
'=====================================================================

Private Const ERR_INVALID_INPUT As Long = vbObjectError + 1001
Private Const ERR_DIV_BY_ZERO    As Long = vbObjectError + 1002
Private Const ERR_NOT_AVAIL      As Long = vbObjectError + 1003
Private Const DBL_INSERTION_THRESHOLD As Long = 16

' Numerical tolerance constants
Private Const NUM_EPS     As Double = 2.22044604925031E-16  ' Machine epsilon
Private Const FPMIN       As Double = 1E-300                ' Floating-point minimum
Private Const TOL_STRICT  As Double = 1E-15                 ' Strict tolerance (SVD/skewness)
Private Const TOL_DEFAULT As Double = 1E-12                 ' Default tolerance (variance/correlation)
Private Const TOL_BISECT  As Double = 1E-14                 ' Bisection convergence

'=============================================================================
'=============================================================================

'=============================================================================
' MeanDouble — 对预提取的 Double() 数组直接计算均值 (避免重复 ExtractDoubles)
'=============================================================================
Private Function MeanDouble(ByRef arr() As Double) As Double
    Dim i As Long, lb As Long
    Dim c As Double, y As Double, t As Double
    Dim sum As Double
    lb = LBound(arr)
    If UBound(arr) < lb Then
        MeanDouble = 0#
        Exit Function
    End If
    For i = lb To UBound(arr)
        y = arr(i) - c
        t = sum + y
        c = (t - sum) - y
        sum = t
    Next i
    MeanDouble = sum / (UBound(arr) - lb + 1)
End Function

'=============================================================================
' StdDevDouble — 对预提取的 Double() 数组直接计算样本标准差
'=============================================================================
Private Function StdDevDouble(ByRef arr() As Double) As Double
    Dim i As Long, n As Long, lb As Long
    Dim m As Double, ss As Double
    Dim c As Double, y As Double, t As Double
    lb = LBound(arr): n = UBound(arr) - lb + 1
    If n < 2 Then
        StdDevDouble = 0#
        Exit Function
    End If
    m = MeanDouble(arr)
    For i = lb To UBound(arr)
        y = (arr(i) - m) ^ 2 - c
        t = ss + y
        c = (t - ss) - y
        ss = t
    Next i
    If ss < 0# Then ss = 0#
    StdDevDouble = Sqr(ss / (n - 1))
End Function

'=============================================================================
' ExtractDoubles — 从 Variant 提取 Double 数组
'
' 支持: Range (单列/单行), 1D Variant 数组, 2D 数组 (可用 colIndex 指定列)
'=============================================================================
Private Function ExtractDoubles( _
    ByRef data As Variant, _
    Optional ByVal colIndex As Long = 1, _
    Optional ByRef outSkippedCount As Variant) As Variant

    Dim result() As Double
    Dim nRows As Long
    Dim i As Long
    Dim cnt As Long
    Dim skipped As Long
    Dim localData As Variant
    Dim arrData As Variant
    Dim rng As Range
    Dim nCols As Long
    Dim lb1 As Long, ub1 As Long
    Dim r0 As Long, c0 As Long, targetCol As Long

    If IsEmpty(data) Then
        ExtractDoubles = Array()
        Exit Function
    End If

    If IsObject(data) And TypeName(data) <> "Range" Then
        If Not IsMissing(outSkippedCount) Then outSkippedCount = -1
        ExtractDoubles = Array()
        Exit Function
    End If

    ' Range → 2D 数组 — 使用本地副本避免修改 ByRef 参数
    If TypeName(data) = "Range" Then
        Set rng = data
        If rng.Count = 0 Then
            ExtractDoubles = Array()
            Exit Function
        End If
        If rng.Count = 1 Then
            If IsNumeric(rng.Value) Then
                ReDim result(0 To 0)
                result(0) = CDbl(rng.Value)
                ExtractDoubles = result
            Else
                If Not IsMissing(outSkippedCount) Then outSkippedCount = 1
                ExtractDoubles = Array()
            End If
            Exit Function
        End If
        localData = rng.Value
    Else
        localData = data
    End If

    If Not IsArray(localData) Then
        If IsNumeric(localData) Then
            ReDim result(0 To 0)
            result(0) = CDbl(localData)
            ExtractDoubles = result
        Else
            If Not IsMissing(outSkippedCount) Then outSkippedCount = 1
            ExtractDoubles = Array()
        End If
        Exit Function
    End If

    ' 检测 1D vs 2D (下标越界 = Error 9 表示无第2维即1D数组)
    arrData = localData
    Err.Clear
    On Error Resume Next
    nCols = UBound(arrData, 2)
    If Err.Number = 9 Then
        ' 1D 数组
        On Error GoTo 0
        lb1 = LBound(arrData): ub1 = UBound(arrData)
        If ub1 < lb1 Then
            ExtractDoubles = Array()
            Exit Function
        End If
        ReDim result(0 To ub1 - lb1)
        cnt = 0
        For i = lb1 To ub1
            If Not IsEmpty(arrData(i)) And VarType(arrData(i)) <> vbBoolean And IsNumeric(arrData(i)) Then
                result(cnt) = CDbl(arrData(i))
                cnt = cnt + 1
            Else
                skipped = skipped + 1
            End If
        Next i
        If Not IsMissing(outSkippedCount) Then outSkippedCount = skipped
        If cnt > 0 Then
            ReDim Preserve result(0 To cnt - 1)
            ExtractDoubles = result
        Else
            ExtractDoubles = Array()
        End If
        Exit Function
    End If
    On Error GoTo 0

    ' 2D 数组
    nRows = UBound(arrData, 1)
    r0 = LBound(arrData, 1)
    If nRows < r0 Then
        ExtractDoubles = Array()
        Exit Function
    End If
    c0 = LBound(arrData, 2)

    If colIndex >= 1 And colIndex <= UBound(arrData, 2) - c0 + 1 Then
        targetCol = c0 + colIndex - 1
    Else
        If Not IsMissing(outSkippedCount) Then outSkippedCount = -1
        ExtractDoubles = Array()
        Exit Function
    End If

    ReDim result(0 To nRows - r0)
    cnt = 0
    For i = r0 To nRows
        If Not IsEmpty(arrData(i, targetCol)) And VarType(arrData(i, targetCol)) <> vbBoolean And IsNumeric(arrData(i, targetCol)) Then
            result(cnt) = CDbl(arrData(i, targetCol))
            cnt = cnt + 1
        Else
            skipped = skipped + 1
        End If
    Next i
    If Not IsMissing(outSkippedCount) Then outSkippedCount = skipped
    If cnt > 0 Then
        ReDim Preserve result(0 To cnt - 1)
        ExtractDoubles = result
    Else
        ExtractDoubles = Array()
    End If
End Function

'=============================================================================
' IsNumericCell — 严格的数值单元格判定 (排除 Empty/Boolean/Date/Error)
'=============================================================================
Private Function IsNumericCell(ByVal v As Variant) As Boolean
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    IsNumericCell = vk.IsNumericCell(v)
End Function

'=============================================================================
'===== 集中趋势 (Central Tendency) — Mean → MeanAbsDev =====
' Mean — 算术平均值
'=============================================================================
Public Function Mean(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "Mean", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    Mean = MeanDouble(arr)
End Function

'=============================================================================
' WeightedMean — 加权算术平均
'=============================================================================
Public Function WeightedMean(ByRef values As Variant, ByRef weights As Variant, _
    Optional ByVal colIdxV As Long = 1, Optional ByVal colIdxW As Long = 1) As Variant
    Dim arrV As Variant, arrW As Variant, sumW As Double, sumWV As Double, i As Long
    Dim cWV As Double, cW As Double, yWV As Double, yW As Double, tWV As Double, tW As Double
    arrV = ExtractDoubles(values, colIdxV): arrW = ExtractDoubles(weights, colIdxW)
    If UBound(arrV) - LBound(arrV) < 0 Or UBound(arrW) - LBound(arrW) <> UBound(arrV) - LBound(arrV) Then
        Err.Raise ERR_INVALID_INPUT, "WeightedMean", "需要至少一个有效数值。": Exit Function
    End If
    Dim v() As Double, w() As Double: v = arrV: w = arrW
    Dim wOffset As Long: wOffset = LBound(arrW) - LBound(arrV)
    For i = LBound(v) To UBound(v)
        yWV = v(i) * w(i + wOffset) - cWV: tWV = sumWV + yWV: cWV = (tWV - sumWV) - yWV: sumWV = tWV
        yW = w(i + wOffset) - cW: tW = sumW + yW: cW = (tW - sumW) - yW: sumW = tW
    Next i
    If sumW = 0# Then Err.Raise ERR_DIV_BY_ZERO, "WeightedMean", "权重之和为零。" Else WeightedMean = sumWV / sumW
End Function

'=============================================================================
' Median — 中位数
'=============================================================================
Public Function Median(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    arrV = ExtractDoubles(data, colIndex)

    If UBound(arrV) - LBound(arrV) + 1 = 0 Then
        Err.Raise ERR_INVALID_INPUT, "Median", "需要至少一个有效数值。"
        Exit Function
    End If

    Dim arr() As Double
    arr = arrV
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1
    QuickSortDouble arr, LBound(arr), UBound(arr)

    Dim mid As Long
    If n Mod 2 = 1 Then
        Median = arr(LBound(arr) + n \ 2)
    Else
        mid = LBound(arr) + n \ 2
        Median = (arr(mid - 1) + arr(mid)) / 2#
    End If
End Function
' Mode — 众数 (返回第一个出现频率最高的值)
'=============================================================================
Public Function Mode(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim dict As Object
    Dim i As Long, lb As Long
    Dim maxCount As Long, modeVal As Double
    Dim currentCount As Long

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 = 0 Then
        Err.Raise ERR_INVALID_INPUT, "Mode", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr)

    Set dict = DP.Create()

    Dim val As Double, key As String, roundedVal As Double
    For i = lb To UBound(arr)
        val = arr(i)
        ' 转换为字符串前先四舍五入到 12 位小数 — 防止
        ' 浮点噪声 (如 1/3 与 2/3-1/3) 导致
        ' 语义相同值产生不同的字典键
        roundedVal = Round(val, 12)
        key = Str$(roundedVal)  ' Str$ 用 "." 小数分隔符, 跨 locale 一致 (#19)
        If dict.Exists(key) Then
            currentCount = CLng(dict(key)) + 1
            dict(key) = currentCount
        Else
            currentCount = 1
            dict.Add key, currentCount
        End If
        If currentCount > maxCount Then
            maxCount = currentCount
            modeVal = roundedVal
        End If
    Next i

    If maxCount <= 1 Then
        Err.Raise ERR_NOT_AVAIL, "Mode", "所有值均为唯一值，无众数。"
    Else
        Mode = modeVal
    End If
End Function

'=============================================================================
' MinMax — 同时返回最小值与最大值
'
'=============================================================================
' Min — 数组最小值 (单值包装 MinMax)
'=============================================================================
Public Function Min(ByRef data As Variant, _
                    Optional ByVal colIndex As Long = 1) As Variant
    Dim outMin As Double, outMax As Double
    Dim result As Variant
    result = MinMax(data, outMin, outMax, colIndex)
    If IsError(result) Then Min = result: Exit Function
    Min = outMin
End Function

'=============================================================================
' Max — 数组最大值 (单值包装 MinMax)
'=============================================================================
Public Function Max(ByRef data As Variant, _
                    Optional ByVal colIndex As Long = 1) As Variant
    Dim outMin As Double, outMax As Double
    Dim result As Variant
    result = MinMax(data, outMin, outMax, colIndex)
    If IsError(result) Then Max = result: Exit Function
    Max = outMax
End Function

'=============================================================================
' MinMax — 同时返回最小值和最大值
' 返回值: Variant 数组 {min, max} (需作为数组公式输入)
' 或使用 ByRef outMin / outMax 获取值 (VBA 调用)
'=============================================================================
Public Function MinMax(ByRef data As Variant, _
                       Optional ByRef outMin As Double = 0#, _
                       Optional ByRef outMax As Double = 0#, _
                       Optional ByVal colIndex As Long = 1) As Variant

    Dim arrV As Variant
    Dim i As Long, lb As Long, n As Long

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 = 0 Then
        Err.Raise ERR_INVALID_INPUT, "MinMax", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    outMin = arr(lb): outMax = arr(lb)
    For i = lb + 1 To UBound(arr)
        If arr(i) < outMin Then outMin = arr(i)
        If arr(i) > outMax Then outMax = arr(i)
    Next i
    MinMax = Array(outMin, outMax)
End Function
'=============================================================================
' GeometricMean — 几何平均数
'
' 适用: 增长率、收益率、比例等乘法场景
'=============================================================================
Public Function GeometricMean(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim i As Long, lb As Long, n As Long

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "GeometricMean", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    ' 使用对数空间计算，避免乘积溢出/下溢 (Kahan 补偿)
    Dim logSum As Double, kC As Double, kY As Double, kT As Double
    logSum = 0#: kC = 0#
    For i = lb To UBound(arr)
        If arr(i) <= 0# Then
            Err.Raise ERR_INVALID_INPUT, "GeometricMean", "所有值必须为正数。"
            Exit Function
        End If
        kY = Log(arr(i)) - kC: kT = logSum + kY: kC = (kT - logSum) - kY: logSum = kT
    Next i
    GeometricMean = Exp(logSum / CDbl(n))
End Function

'=============================================================================
' HarmonicMean — 调和平均数
'
' 适用: 速度、比率等倒数场景
'=============================================================================
Public Function HarmonicMean(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim i As Long, lb As Long, n As Long
    Dim sumRecip As Double
    Dim kC As Double, kY As Double, kT As Double  ' Kahan compensation vars
    kC = 0#

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "HarmonicMean", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    For i = lb To UBound(arr)
        If arr(i) <= 0# Then
            Err.Raise ERR_INVALID_INPUT, "HarmonicMean", "所有值必须为正数。"
            Exit Function
        End If
        ' Kahan 补偿求和 (与模块其他求和路径一致)
        kY = 1# / arr(i) - kC: kT = sumRecip + kY: kC = (kT - sumRecip) - kY: sumRecip = kT
    Next i
    HarmonicMean = n / sumRecip
End Function

'=============================================================================
' TrimMean — 截尾均值
'
' 剔除两端各 trimPct/2 的数据后计算均值
' 例: trimPct=0.1 剔除 5% 最小和 5% 最大
'=============================================================================
Public Function TrimMean( _
    ByRef data As Variant, _
    Optional ByVal trimPct As Double = 0.1, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim arrV As Variant
    Dim lb As Long, n As Long
    Dim trimN As Long, i As Long
    Dim total As Double, cnt As Long

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "TrimMean", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1
    If n < 3 Then
        TrimMean = MeanDouble(arr)
        Exit Function
    End If
    If trimPct < 0# Then trimPct = 0#
    If trimPct >= 1# Then trimPct = 0.5

    QuickSortDouble arr, lb, UBound(arr)
    trimN = Int(n * trimPct / 2#)
    If trimN < 0 Then trimN = 0
    If trimN * 2 >= n Then trimN = (n - 1) \ 2

    Dim kC2 As Double, kY2 As Double, kT2 As Double
    For i = lb + trimN To UBound(arr) - trimN
        kY2 = arr(i) - kC2: kT2 = total + kY2: kC2 = (kT2 - total) - kY2: total = kT2
        cnt = cnt + 1
    Next i
    If cnt > 0 Then TrimMean = total / cnt Else Err.Raise ERR_INVALID_INPUT, "TrimMean", "需要至少一个有效数值。"
End Function

'=============================================================================
' RootMeanSquare — 均方根 RMS
'=============================================================================
Public Function RootMeanSquare(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim i As Long, lb As Long, n As Long
    Dim ss As Double, c As Double, y As Double, t As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "RootMeanSquare", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    For i = lb To UBound(arr)
        y = arr(i) ^ 2 - c: t = ss + y: c = (t - ss) - y: ss = t
    Next i
    RootMeanSquare = Sqr(ss / n)
End Function

'=============================================================================
' MeanAbsDev — 平均绝对偏差
'=============================================================================
Public Function MeanAbsDev(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim i As Long, lb As Long, n As Long
    Dim m As Double, totalDev As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "MeanAbsDev", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    m = MeanDouble(arr)
    Dim kC3 As Double, kY3 As Double, kT3 As Double
    For i = lb To UBound(arr)
        kY3 = Abs(arr(i) - m) - kC3: kT3 = totalDev + kY3: kC3 = (kT3 - totalDev) - kY3: totalDev = kT3
    Next i
    MeanAbsDev = totalDev / n
End Function


'=============================================================================
'===== 离散度 (Dispersion) — StdDev → IQR =====
' StdDev / StdDevP — 样本标准差 / 总体标准差
'=============================================================================
Public Function StdDev(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim v As Variant
    v = Variance(data, colIndex)
    If IsError(v) Then
        StdDev = v
    Else
        StdDev = Sqr(v)
    End If
End Function

Public Function StdDevP(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim v As Variant
    v = VarianceP(data, colIndex)
    If IsError(v) Then
        StdDevP = v
    Else
        StdDevP = Sqr(v)
    End If
End Function

'=============================================================================
' Variance / VarianceP — 样本方差(n-1) / 总体方差(n)
'=============================================================================
Public Function Variance(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim i As Long, n As Long, lb As Long
    Dim m As Double, ss As Double
    Dim c As Double, y As Double, t As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 2 Then
        Err.Raise ERR_INVALID_INPUT, "Variance", "需要至少 2 个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    m = MeanDouble(arr)
    For i = lb To UBound(arr)
        y = (arr(i) - m) ^ 2 - c
        t = ss + y
        c = (t - ss) - y
        ss = t
    Next i
    Variance = ss / (n - 1)
    If Variance < 0# Then Variance = 0#
End Function

Public Function VarianceP(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim i As Long, n As Long, lb As Long
    Dim m As Double, ss As Double
    Dim c As Double, y As Double, t As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "VarianceP", "需要至少 1 个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    m = MeanDouble(arr)
    For i = lb To UBound(arr)
        y = (arr(i) - m) ^ 2 - c
        t = ss + y
        c = (t - ss) - y
        ss = t
    Next i
    VarianceP = ss / n
    If VarianceP < 0# Then VarianceP = 0#
End Function

'=============================================================================
' Percentile — 百分位数 (线性插值法, 0 ≤ k ≤ 1)
'=============================================================================
Public Function Percentile( _
    ByRef data As Variant, _
    ByVal k As Double, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim arrV As Variant
    Dim lb As Long, ub As Long, n As Long

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 = 0 Then
        Err.Raise ERR_INVALID_INPUT, "Percentile", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1
    If k < 0# Or k > 1# Then
        Err.Raise ERR_INVALID_INPUT, "Percentile", "分位值必须在 0-1 之间。"
        Exit Function
    End If

    QuickSortDouble arr, lb, UBound(arr)
    Percentile = QuickPercentile(arr, k, lb, UBound(arr))
End Function

'=============================================================================
' IQR — 四分位距 = Q3 - Q1
'=============================================================================
Public Function IQR(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim lb As Long, ub As Long, n As Long

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 = 0 Then
        Err.Raise ERR_INVALID_INPUT, "IQR", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): ub = UBound(arr)
    n = ub - lb + 1

    QuickSortDouble arr, lb, ub
    IQR = QuickPercentile(arr, 0.75, lb, ub) - QuickPercentile(arr, 0.25, lb, ub)
End Function

Private Function QuickPercentile(ByRef arr() As Double, ByVal k As Double, ByVal lb As Long, ByVal ub As Long) As Double
    Dim n As Long: n = ub - lb + 1
    Dim idx As Double: idx = k * (n - 1)
    Dim lo As Long: lo = lb + Int(idx)
    Dim hi As Long: hi = lo + 1
    If hi > ub Then hi = ub
    Dim frac As Double: frac = idx - Int(idx)
    QuickPercentile = arr(lo) + frac * (arr(hi) - arr(lo))
End Function

'===== 形态 (Shape) — Skewness → Kurtosis =====
' Skewness — 样本偏度 (adjusted Fisher-Pearson)
'
' 正值 = 右偏, 负值 = 左偏, 接近 0 = 对称
'=============================================================================
Public Function Skewness(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim i As Long, n As Long, lb As Long
    Dim m As Double, s As Double, sum3 As Double
    Dim c As Double, y As Double, t As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 3 Then
        Err.Raise ERR_INVALID_INPUT, "Skewness", "需要至少 3 个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    m = MeanDouble(arr)
    s = StdDevDouble(arr)
    If Abs(s) < TOL_STRICT Then
        Err.Raise ERR_DIV_BY_ZERO, "Skewness", "标准差为零，无法计算偏度。"
        Exit Function
    End If

    For i = lb To UBound(arr)
        y = ((arr(i) - m) / s) ^ 3 - c
        t = sum3 + y
        c = (t - sum3) - y
        sum3 = t
    Next i
    Skewness = (CDbl(n) / (CDbl(n - 1) * CDbl(n - 2))) * sum3
End Function

'=============================================================================
' Kurtosis — 超额峰度 (样本)
'
' >0 = 尖峰厚尾, <0 = 扁平瘦尾, ~0 = 正态分布
'=============================================================================
Public Function Kurtosis(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    Dim i As Long, n As Long, lb As Long
    Dim m As Double, s As Double, sum4 As Double
    Dim c As Double, y As Double, t As Double
    Dim term1 As Double, term2 As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 4 Then
        Err.Raise ERR_INVALID_INPUT, "Kurtosis", "需要至少 4 个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    m = MeanDouble(arr)
    s = StdDevDouble(arr)
    If Abs(s) < TOL_STRICT Then
        Err.Raise ERR_DIV_BY_ZERO, "Kurtosis", "标准差为零，无法计算峰度。"
        Exit Function
    End If

    For i = lb To UBound(arr)
        y = ((arr(i) - m) / s) ^ 4 - c
        t = sum4 + y
        c = (t - sum4) - y
        sum4 = t
    Next i

    term1 = (CDbl(n) * (n + 1)) / (CDbl(n - 1) * CDbl(n - 2) * CDbl(n - 3))
    term2 = 3# * (n - 1) ^ 2 / (CDbl(n - 2) * CDbl(n - 3))
    Kurtosis = term1 * sum4 - term2
End Function
'===== 排名 (Ranking) — Rank → PercentRank =====
' Rank — 值在数据集中的排名 (1 = 最小或最大)
'
' ascending=True:  1 = 最小值 (排名升序)
' ascending=False: 1 = 最大值 (排名降序)
' 注意: 使用改良竞争排名 — ties are counted as ≤ or ≥ the value,
'       不返回小数排名。例如 [1,2,2,3] 中值为 2 的排名返回 3 (升序)
'=============================================================================
Public Function Rank( _
    ByRef data As Variant, _
    ByRef value As Variant, _
    Optional ByVal ascending As Boolean = True, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim arrV As Variant
    Dim i As Long, lb As Long, n As Long, cnt As Long

    If VarType(value) = vbBoolean Or Not IsNumeric(value) Then
        Err.Raise ERR_INVALID_INPUT, "Rank", "需要至少一个有效数值。"
        Exit Function
    End If

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "Rank", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): n = UBound(arr) - lb + 1

    Dim val As Double: val = CDbl(value)

    If ascending Then
        For i = lb To UBound(arr)
            If arr(i) <= val Then cnt = cnt + 1
        Next i
    Else
        For i = lb To UBound(arr)
            If arr(i) >= val Then cnt = cnt + 1
        Next i
    End If
    Rank = cnt
End Function

'=============================================================================
' RankEq — 与 Excel RANK.EQ 一致: ties 同排名，后续跳过
'=============================================================================
Public Function RankEq(ByRef data As Variant, ByRef value As Variant, _
    Optional ByVal ascending As Boolean = True, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant: If VarType(value) = vbBoolean Or Not IsNumeric(value) Then Err.Raise ERR_INVALID_INPUT, "RankEq", "需要至少一个有效数值。": Exit Function
    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then Err.Raise ERR_INVALID_INPUT, "RankEq", "需要至少一个有效数值。": Exit Function
    Dim arr() As Double, val As Double: arr = arrV: val = CDbl(value)
    Dim i As Long, cnt As Long, lb As Long: lb = LBound(arr)
    If ascending Then
        For i = lb To UBound(arr)
            If arr(i) < val Then cnt = cnt + 1
        Next i
    Else
        For i = lb To UBound(arr)
            If arr(i) > val Then cnt = cnt + 1
        Next i
    End If
    RankEq = cnt + 1
End Function

'=============================================================================
' RankAvg — 与 Excel RANK.AVG 一致: ties 取平均排名
'=============================================================================
Public Function RankAvg(ByRef data As Variant, ByRef value As Variant, _
    Optional ByVal ascending As Boolean = True, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant: If VarType(value) = vbBoolean Or Not IsNumeric(value) Then Err.Raise ERR_INVALID_INPUT, "RankAvg", "需要至少一个有效数值。": Exit Function
    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then Err.Raise ERR_INVALID_INPUT, "RankAvg", "需要至少一个有效数值。": Exit Function
    Dim arr() As Double, val As Double: arr = arrV: val = CDbl(value)
    Dim i As Long, lt As Long, eq As Long, lb As Long: lb = LBound(arr)
    If ascending Then
        For i = lb To UBound(arr)
            If arr(i) < val Then
                lt = lt + 1
            ElseIf Abs(arr(i) - val) <= TOL_DEFAULT * Application.Max(Abs(arr(i)), Abs(val), 1#) Then
                eq = eq + 1
            End If
        Next i
    Else
        For i = lb To UBound(arr)
            If arr(i) > val Then
                lt = lt + 1
            ElseIf Abs(arr(i) - val) <= TOL_DEFAULT * Application.Max(Abs(arr(i)), Abs(val), 1#) Then
                eq = eq + 1
            End If
        Next i
    End If
    If eq = 0 Then Err.Raise ERR_NOT_AVAIL, "RankAvg", "数值不在数据范围内。" Else RankAvg = lt + (eq + 1#) / 2#
End Function

'=============================================================================
' PercentRank — 百分位排名 (0-1)
'=============================================================================
Public Function PercentRank( _
    ByRef data As Variant, _
    ByRef value As Variant, _
    Optional ByVal ascending As Boolean = True, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim arrV As Variant
    Dim n As Long

    If VarType(value) = vbBoolean Or Not IsNumeric(value) Then
        Err.Raise ERR_INVALID_INPUT, "PercentRank", "需要至少一个有效数值。"
        Exit Function
    End If

    arrV = ExtractDoubles(data, colIndex)
    n = UBound(arrV) - LBound(arrV) + 1
    If n < 2 Then
        Err.Raise ERR_INVALID_INPUT, "PercentRank", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV

    Dim val As Double: val = CDbl(value)
    Dim r As Long, i As Long, lb As Long
    lb = LBound(arr): r = 0
    If ascending Then
        For i = lb To UBound(arr)
            If arr(i) <= val Then r = r + 1
        Next i
    Else
        For i = lb To UBound(arr)
            If arr(i) >= val Then r = r + 1
        Next i
    End If

    If r = 0 Then
        Err.Raise ERR_INVALID_INPUT, "PercentRank", "需要至少一个有效数值。"
    Else
        PercentRank = (r - 1#) / (n - 1#)
    End If
End Function

'=============================================================================

'===== 关联分析 (Association) — Covariance → CorrelationMatrix =====
' ExtractPairedDoubles — 成对提取两个数据源的数值行
'
' 与分别调用 ExtractDoubles 不同, 此函数确保两列的行一一对应:
'   仅保留两列均为数值的行, 避免非数值行位置不同导致的错配。
'=============================================================================
Private Sub ExtractPairedDoubles( _
    ByRef dataX As Variant, _
    ByRef dataY As Variant, _
    ByRef outArrX As Variant, _
    ByRef outArrY As Variant, _
    ByVal colX As Long, _
    ByVal colY As Long)

    Dim arrXRaw As Variant, arrYRaw As Variant
    Dim nRows As Long, r As Long, lbR As Long, ubR As Long
    Dim cnt As Long, i As Long, yLb As Long, yCount As Long
    Dim cX As Long, cY As Long
    Dim nx As Long, ny As Long
    Dim tmpX() As Double, tmpY() As Double

    ' 将 Range 转换为 2D 数组
    If TypeName(dataX) = "Range" Then arrXRaw = dataX.Value Else arrXRaw = dataX
    If TypeName(dataY) = "Range" Then arrYRaw = dataY.Value Else arrYRaw = dataY

    ' 两者必须为 2D 数组才能进行成对提取
    If Not IsArray(arrXRaw) Or Not IsArray(arrYRaw) Then GoTo Fallback

    Err.Clear
    On Error Resume Next
    ubR = UBound(arrXRaw, 1)
    If Err.Number <> 0 Then GoTo Fallback
    If UBound(arrYRaw, 1) <> ubR Then GoTo Fallback

    ' 确定列索引 (保持 On Error Resume Next 以安全处理 1D 情况)
    lbR = LBound(arrXRaw, 1)
    nRows = ubR - lbR + 1
    cX = LBound(arrXRaw, 2) + colX - 1
    If Err.Number <> 0 Then GoTo Fallback
    cY = LBound(arrYRaw, 2) + colY - 1
    If Err.Number <> 0 Then GoTo Fallback
    If cX > UBound(arrXRaw, 2) Or cY > UBound(arrYRaw, 2) Then GoTo Fallback
    On Error GoTo 0

    ' 成对提取: 仅保留两列均为数值的行
    ReDim tmpX(0 To nRows - 1)
    ReDim tmpY(0 To nRows - 1)
    cnt = 0
    For r = lbR To ubR
        If Not IsEmpty(arrXRaw(r, cX)) And Not IsEmpty(arrYRaw(r, cY)) And _
           VarType(arrXRaw(r, cX)) <> vbBoolean And VarType(arrYRaw(r, cY)) <> vbBoolean And _
           IsNumeric(arrXRaw(r, cX)) And IsNumeric(arrYRaw(r, cY)) Then
            tmpX(cnt) = CDbl(arrXRaw(r, cX))
            tmpY(cnt) = CDbl(arrYRaw(r, cY))
            cnt = cnt + 1
        End If
    Next r
    If cnt > 0 Then
        ReDim Preserve tmpX(0 To cnt - 1)
        ReDim Preserve tmpY(0 To cnt - 1)
        outArrX = tmpX
        outArrY = tmpY
    Else
        outArrX = Array()
        outArrY = Array()
    End If
    Exit Sub

Fallback:
    On Error GoTo 0
    
    ' 尝试 1D 数组的对齐提取 (关键修复: 避免独立调用
    ' ExtractDoubles 分别删除非数值行,
    ' 导致配对错位)
    If IsArray(arrXRaw) And IsArray(arrYRaw) Then
        Err.Clear
        On Error Resume Next
        ubR = UBound(arrXRaw, 2)
        If Err.Number = 9 Then
            Err.Clear
            ' 两者均为 1D 数组 — 按索引对齐进行成对提取
            lbR = LBound(arrXRaw): ubR = UBound(arrXRaw)
            yLb = LBound(arrYRaw): yCount = UBound(arrYRaw) - yLb + 1
            nRows = ubR - lbR + 1
            If nRows > yCount Then nRows = yCount
            On Error GoTo 0
            If nRows < 1 Then outArrX = Array(): outArrY = Array(): Exit Sub
            ReDim tmpX(0 To nRows - 1)
            ReDim tmpY(0 To nRows - 1)
            cnt = 0
            For i = 0 To nRows - 1
                If VarType(arrXRaw(lbR + i)) <> vbBoolean And VarType(arrYRaw(yLb + i)) <> vbBoolean And _
                   IsNumeric(arrXRaw(lbR + i)) And IsNumeric(arrYRaw(yLb + i)) Then
                    tmpX(cnt) = CDbl(arrXRaw(lbR + i))
                    tmpY(cnt) = CDbl(arrYRaw(yLb + i))
                    cnt = cnt + 1
                End If
            Next i
            If cnt > 0 Then
                ReDim Preserve tmpX(0 To cnt - 1)
                ReDim Preserve tmpY(0 To cnt - 1)
                outArrX = tmpX
                outArrY = tmpY
            Else
                outArrX = Array()
                outArrY = Array()
            End If
            Exit Sub
        End If
        On Error GoTo 0
    End If
    
    ' 无法对齐的行配对输入 — 返回空数组以避免静默错误结果
    ' ExtractDoubles 独立过滤会丢弃不同位置的非数值行导致错配。
    ' 到达此处说明输入既非正常 2D 数组也非可对齐的 1D 数组。
    outArrX = Array()
    outArrY = Array()
End Sub

'=============================================================================
' CovarianceDouble — 对预提取的 Double() 数组计算样本协方差
'=============================================================================
Private Function CovarianceDouble(ByRef arrX() As Double, ByRef arrY() As Double) As Double
    Dim i As Long, n As Long, lb As Long
    Dim mx As Double, my As Double, ss As Double
    Dim c As Double, y As Double, t As Double  ' Kahan 补偿求和变量

    lb = LBound(arrX)
    n = UBound(arrX) - lb + 1
    If n < 2 Or UBound(arrY) - LBound(arrY) + 1 <> n Then
        CovarianceDouble = 0#
        Exit Function
    End If

    mx = MeanDouble(arrX)
    my = MeanDouble(arrY)
    ' 处理 arrX 与 arrY 可能不等的 LBound
    Dim offset As Long: offset = LBound(arrY) - lb
    For i = lb To UBound(arrX)
        y = (arrX(i) - mx) * (arrY(i + offset) - my) - c
        t = ss + y
        c = (t - ss) - y
        ss = t
    Next i
    CovarianceDouble = ss / (n - 1)
End Function

'=============================================================================
' Covariance — 样本协方差
'=============================================================================
Public Function Covariance( _
    ByRef dataX As Variant, _
    ByRef dataY As Variant, _
    Optional ByVal colX As Long = 1, _
    Optional ByVal colY As Long = 1) As Variant

    Dim arrX As Variant, arrY As Variant
    ExtractPairedDoubles dataX, dataY, arrX, arrY, colX, colY

    If UBound(arrX) - LBound(arrX) + 1 < 2 Then
        Err.Raise ERR_INVALID_INPUT, "Covariance", "需要至少 2 对有效数值。"
        Exit Function
    End If
    Dim arrXD() As Double, arrYD() As Double
    arrXD = arrX: arrYD = arrY

    Covariance = CovarianceDouble(arrXD, arrYD)
End Function

'=============================================================================
' Correlation — Pearson 相关系数
'=============================================================================
Public Function Correlation( _
    ByRef dataX As Variant, _
    ByRef dataY As Variant, _
    Optional ByVal colX As Long = 1, _
    Optional ByVal colY As Long = 1) As Variant

    Dim arrX As Variant, arrY As Variant
    Dim n As Long, lb As Long
    Dim sx As Double, sy As Double

    ExtractPairedDoubles dataX, dataY, arrX, arrY, colX, colY

    If UBound(arrX) - LBound(arrX) + 1 < 2 Then
        Err.Raise ERR_INVALID_INPUT, "Correlation", "需要至少 2 对有效数值。"
        Exit Function
    End If
    Dim arrXD() As Double, arrYD() As Double
    arrXD = arrX: arrYD = arrY
    lb = LBound(arrXD)
    n = UBound(arrXD) - lb + 1

    sx = StdDevDouble(arrXD)
    sy = StdDevDouble(arrYD)
    ' 尺度感知容差: 相对数据最大绝对值, 避免误杀小幅数据 (如 1e-14 尺度)
    If sx <= TOL_DEFAULT * MaxAbsDouble(arrXD) Or sy <= TOL_DEFAULT * MaxAbsDouble(arrYD) Then
        Err.Raise ERR_DIV_BY_ZERO, "Correlation", "方差为零，无法计算相关系数。"
        Exit Function
    End If

    Correlation = CovarianceDouble(arrXD, arrYD) / (sx * sy)
End Function
' RSquare — R² 判定系数
'=============================================================================
Public Function RSquare( _
    ByRef actual As Variant, _
    ByRef predicted As Variant, _
    Optional ByVal colActual As Long = 1, _
    Optional ByVal colPredicted As Long = 1) As Variant

    Dim arrA As Variant, arrP As Variant
    Dim i As Long, n As Long, lb As Long
    Dim meanA As Double, ssTotal As Double, ssResid As Double
    Dim cT As Double, cR As Double, y As Double, t As Double

    ExtractPairedDoubles actual, predicted, arrA, arrP, colActual, colPredicted

    If UBound(arrA) - LBound(arrA) + 1 < 2 Then
        Err.Raise ERR_INVALID_INPUT, "RSquare", "需要至少 2 对有效数值。"
        Exit Function
    End If
    Dim arrAD() As Double, arrPD() As Double
    arrAD = arrA: arrPD = arrP
    lb = LBound(arrAD): n = UBound(arrAD) - lb + 1

    meanA = MeanDouble(arrAD)
    For i = lb To UBound(arrAD)
        y = (arrAD(i) - meanA) ^ 2 - cT
        t = ssTotal + y
        cT = (t - ssTotal) - y
        ssTotal = t

        y = (arrAD(i) - arrPD(LBound(arrPD) + i - lb)) ^ 2 - cR
        t = ssResid + y
        cR = (t - ssResid) - y
        ssResid = t
    Next i

    If ssTotal <= TOL_DEFAULT * CDbl(n) * MaxAbsDouble(arrAD) * MaxAbsDouble(arrAD) Then
        Err.Raise ERR_DIV_BY_ZERO, "RSquare", "SS_total 为零，无法计算 R²。"
    Else
        RSquare = 1# - ssResid / ssTotal
    End If
End Function
'=============================================================================
' CorrelationMatrix — 计算数值列之间的 Pearson 相关系数矩阵
'
' 参数:
'   data — Range 或 2D 数组 (首行为表头)
'
' 返回: 2D Variant 数组 (含标签)，第 1 行/列为列名，其余为相关系数
'=============================================================================
Public Function CorrelationMatrix(ByVal data As Variant, _
                                   Optional ByVal hasHeader As Boolean = True) As Variant()
    Dim dataArr As Variant
    Dim numRows As Long, numCols As Long
    Dim minRows As Long
    Dim err1() As Variant, err2() As Variant
    Dim colNames() As String
    Dim numColMap As Variant
    Dim nc As Long
    Dim rowOff As Long
    Dim n As Long
    Dim result() As Variant
    Dim i As Long, j As Long
    Dim colValid() As Variant
    Dim c As Long, c1 As Long, c2 As Long
    Dim mask() As Boolean
    Dim hasValid As Boolean
    Dim colIdx1 As Long, colIdx2 As Long
    Dim nPair As Long
    Dim sum1 As Double, sum2 As Double
    Dim mean1 As Double, mean2 As Double
    Dim ss1 As Double, ss2 As Double, cov As Double
    Dim k_s1 As Double, k_s2 As Double, k_ss1 As Double, k_ss2 As Double, k_cov As Double
    Dim ky As Double, kt As Double
    Dim std1 As Double, std2 As Double
    Dim v1 As Double, v2 As Double
    Dim mask1() As Boolean, m2() As Boolean
    Dim vk As VariantKit: Set vk = New VariantKit
    Dim ao As ArrayOps: Set ao = New ArrayOps

    dataArr = vk.NormalizeTo2D(data, numRows, numCols)
    If hasHeader Then minRows = 3 Else minRows = 2
    If numRows < minRows Then
        ReDim err1(1 To 1, 1 To 1): err1(1, 1) = "Need >=2 data rows": CorrelationMatrix = err1: Exit Function
    End If

    ' 收集数值列
    numColMap = ao.CollectNumericColumns(dataArr, numRows, numCols, colNames, hasHeader, vk)
    If UBound(numColMap) < LBound(numColMap) Then
        ReDim err2(1 To 1, 1 To 1): err2(1, 1) = "No numeric columns": CorrelationMatrix = err2: Exit Function
    End If
    nc = UBound(numColMap)

    ' 构建结果数组 (行列标签)
    If hasHeader Then rowOff = 1 Else rowOff = 0
    n = numRows - rowOff
    ReDim result(1 To nc + 1, 1 To nc + 1)
    result(1, 1) = "Correlation"
    For j = 1 To nc
        result(1, j + 1) = colNames(numColMap(j))
        result(j + 1, 1) = colNames(numColMap(j))
    Next

    ' 预计算每列的有效性掩码, 避免在内层循环中重复调用 IsNumericCell
    ReDim colValid(1 To nc)
    For c = 1 To nc
        ReDim mask(1 To n)
        hasValid = False
        For i = 1 To n
            If IsNumericCell(dataArr(i + rowOff, numColMap(c))) Then
                mask(i) = True: hasValid = True
            End If
        Next
        If hasValid Then colValid(c) = mask
    Next

    ' 计算相关系数矩阵 — 成对删除缺失值 (pairwise complete observations)
    For c1 = 1 To nc
        colIdx1 = numColMap(c1)
        ' 提前提取 c1 的掩码, 避免在内层 c2 循环中重复解包
        If Not IsEmpty(colValid(c1)) Then mask1 = colValid(c1)
        For c2 = 1 To nc
            If c1 = c2 Then
                result(c1 + 1, c2 + 1) = 1#
            Else
                colIdx2 = numColMap(c2)
                ' 第一遍: 使用两列同时有效的行计算均值
                nPair = 0
                sum1 = 0#: sum2 = 0#: k_s1 = 0#: k_s2 = 0#
                If Not IsEmpty(colValid(c1)) And Not IsEmpty(colValid(c2)) Then
                    m2 = colValid(c2)
                    For i = 1 To n
                        If mask1(i) Then
                            If m2(i) Then
                                nPair = nPair + 1
                                ky = CDbl(dataArr(i + rowOff, colIdx1)) - k_s1: kt = sum1 + ky: k_s1 = (kt - sum1) - ky: sum1 = kt
                                ky = CDbl(dataArr(i + rowOff, colIdx2)) - k_s2: kt = sum2 + ky: k_s2 = (kt - sum2) - ky: sum2 = kt
                            End If
                        End If
                    Next
                End If

                If nPair < 2 Then
                    ' 有效样本量不足 — 返回 Empty (nPair=1 无法计算方差)
                    result(c1 + 1, c2 + 1) = Empty
                Else
                    mean1 = sum1 / nPair
                    mean2 = sum2 / nPair

                    ' 第二遍: 计算协方差与标准差
                    ss1 = 0#: ss2 = 0#: cov = 0#: k_ss1 = 0#: k_ss2 = 0#: k_cov = 0#
                    For i = 1 To n
                        If mask1(i) Then
                            If m2(i) Then
                                v1 = CDbl(dataArr(i + rowOff, colIdx1)) - mean1
                                v2 = CDbl(dataArr(i + rowOff, colIdx2)) - mean2
                                ky = v1 * v1 - k_ss1: kt = ss1 + ky: k_ss1 = (kt - ss1) - ky: ss1 = kt
                                ky = v2 * v2 - k_ss2: kt = ss2 + ky: k_ss2 = (kt - ss2) - ky: ss2 = kt
                                ky = v1 * v2 - k_cov: kt = cov + ky: k_cov = (kt - cov) - ky: cov = kt
                            End If
                        End If
                    Next

                    If ss1 > 0# And ss2 > 0# Then
                        std1 = Sqr(ss1 / (nPair - 1))
                        std2 = Sqr(ss2 / (nPair - 1))
                        result(c1 + 1, c2 + 1) = cov / ((nPair - 1) * std1 * std2)
                    Else
                        result(c1 + 1, c2 + 1) = 0#
                    End If
                End If
            End If
        Next
    Next

    CorrelationMatrix = result
End Function

'=============================================================================

'=============================================================================

'=============================================================================
'===== 数据变换 (Transforms) — ZScore → Binning =====
' ZScore — Z 标准化, 可返回单个值或数组
'
' 若 value 非空则返回 (value - mean) / stddev
' 否则返回包含所有 Z 值的数组
'=============================================================================
Public Function ZScore( _
    ByRef data As Variant, _
    Optional ByVal value As Variant, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim arrV As Variant, m As Double, s As Double
    Dim i As Long, lb As Long
    Dim result() As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 2 Then
        Err.Raise ERR_INVALID_INPUT, "ZScore", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr)

    m = MeanDouble(arr)
    s = StdDevDouble(arr)
    ' 尺度感知容差: 相对数据最大绝对值, 避免误杀小幅数据
    If s <= TOL_DEFAULT * MaxAbsDouble(arr) Then
        Err.Raise ERR_DIV_BY_ZERO, "ZScore", "标准差为零，无法计算 Z-Score。"
        Exit Function
    End If

    If Not IsMissing(value) And Not IsEmpty(value) And Not IsError(value) And VarType(value) <> vbBoolean And IsNumeric(value) Then
        ZScore = (CDbl(value) - m) / s
    ElseIf Not IsMissing(value) And Not IsEmpty(value) And IsError(value) Then
        Err.Raise ERR_INVALID_INPUT, "ZScore", "需要至少一个有效数值。"
        Exit Function
    Else
        ReDim result(lb To UBound(arr))
        For i = lb To UBound(arr)
            result(i) = (arr(i) - m) / s
        Next i
        ZScore = result
    End If
End Function

'=============================================================================
' Normalize — Min-Max 归一化到 [0, 1]
'=============================================================================
Public Function Normalize( _
    ByRef data As Variant, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim arrV As Variant
    Dim i As Long, lb As Long
    Dim minVal As Double, maxVal As Double, rng As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 = 0 Then
        Err.Raise ERR_INVALID_INPUT, "Normalize", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr)

    minVal = arr(lb): maxVal = arr(lb)
    For i = lb + 1 To UBound(arr)
        If arr(i) < minVal Then minVal = arr(i)
        If arr(i) > maxVal Then maxVal = arr(i)
    Next i

    rng = maxVal - minVal
    Dim result() As Double
    ReDim result(lb To UBound(arr))

    ' 尺度感知容差: 相对数据最大绝对值, 避免误杀小幅数据
    Dim scaleR As Double: scaleR = Abs(maxVal)
    If Abs(minVal) > scaleR Then scaleR = Abs(minVal)
    If rng <= TOL_DEFAULT * scaleR Then
        Err.Raise ERR_DIV_BY_ZERO, "Normalize", "范围为零，无法归一化。"
        Exit Function
    Else
        For i = lb To UBound(arr)
            result(i) = (arr(i) - minVal) / rng
        Next i
    End If
    Normalize = result
End Function

'=============================================================================
' LinInterp — 线性插值
'
' 参数:
'   x     - 待插值点
'   xs    - X 值序列 (必须已排序)
'   ys    - Y 值序列 (与 xs 长度相同)
'
' 若 x 超出范围, 返回最近边界值 (外推)
'=============================================================================
Public Function LinInterp( _
    ByVal x As Double, _
    ByRef xs As Variant, _
    ByRef ys As Variant) As Variant

    Dim arrVX As Variant, arrVY As Variant
    Dim i As Long, lb As Long, ub As Long
    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double

    arrVX = ExtractDoubles(xs)
    arrVY = ExtractDoubles(ys)

    If UBound(arrVX) - LBound(arrVX) < 0 Or UBound(arrVY) - LBound(arrVY) <> UBound(arrVX) - LBound(arrVX) Then
        Err.Raise ERR_INVALID_INPUT, "LinInterp", "需要至少 2 个点进行插值。"
        Exit Function
    End If
    Dim arrX() As Double, arrY() As Double
    arrX = arrVX: arrY = arrVY
    lb = LBound(arrX): ub = UBound(arrX)

    If x <= arrX(lb) Then
        LinInterp = arrY(lb)
        Exit Function
    End If
    If x >= arrX(ub) Then
        LinInterp = arrY(ub)
        Exit Function
    End If

    Dim lo As Long, hi As Long, mid As Long
    lo = lb: hi = ub
    Do While lo < hi
        mid = (lo + hi) \ 2
        If arrX(mid) < x Then lo = mid + 1 Else hi = mid
    Loop
    If lo > ub Then lo = ub
    If lo <= lb Then lo = lb + 1

    x1 = arrX(lo - 1): x2 = arrX(lo)
    y1 = arrY(lo - 1): y2 = arrY(lo)
    If x2 = x1 Then
        LinInterp = (y1 + y2) / 2#
    Else
        LinInterp = y1 + (y2 - y1) * (x - x1) / (x2 - x1)
    End If
End Function
' Winsorize — 缩尾处理
'
' 将两端各 pct/2 的数据压缩到边界值
'=============================================================================
Public Function Winsorize( _
    ByRef data As Variant, _
    Optional ByVal pct As Double = 0.05, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim arrV As Variant
    Dim result() As Double
    Dim i As Long, lb As Long, ub As Long, n As Long
    Dim trimN As Long
    Dim lowerVal As Double, upperVal As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "Winsorize", "需要至少一个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): ub = UBound(arr)
    n = ub - lb + 1
    If n < 3 Then
        ' 直接复制
        ReDim result(lb To ub)
        For i = lb To ub: result(i) = arr(i): Next i
        Winsorize = result
        Exit Function
    End If

    If pct < 0# Then pct = 0#
    If pct >= 1# Then pct = 0.5

    ' 在副本上排序以确定边界值，保持原始顺序
    Dim sorted() As Double
    ReDim sorted(lb To ub)
    For i = lb To ub: sorted(i) = arr(i): Next i
    QuickSortDouble sorted, lb, ub

    trimN = Int(n * pct / 2#)
    If trimN < 0 Then trimN = 0
    If trimN * 2 >= n Then trimN = (n - 1) \ 2

    lowerVal = sorted(lb + trimN)
    upperVal = sorted(ub - trimN)

    ReDim result(lb To ub)
    For i = lb To ub
        If arr(i) < lowerVal Then
            result(i) = lowerVal
        ElseIf arr(i) > upperVal Then
            result(i) = upperVal
        Else
            result(i) = arr(i)
        End If
    Next i
    Winsorize = result
End Function

'=============================================================================
' MovingAverage — 简单移动平均
'=============================================================================
Public Function MovingAverage( _
    ByRef data As Variant, _
    ByVal window As Long, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim arrV As Variant
    Dim result() As Double
    Dim i As Long, lb As Long, ub As Long, n As Long
    Dim total As Double

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Then
        Err.Raise ERR_INVALID_INPUT, "MovingAverage", "窗口大小或数据不足。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): ub = UBound(arr)
    n = ub - lb + 1

    If window < 1 Then
        Err.Raise ERR_INVALID_INPUT, "MovingAverage", "窗口大小或数据不足。"
        Exit Function
    End If

    If window > n Then window = n
    ReDim result(lb To ub)

    total = 0#
    For i = lb To ub
        total = total + arr(i)
        If i - window >= lb Then total = total - arr(i - window)
        If i >= lb + window - 1 Then
            result(i) = total / window
        Else
            result(i) = total / (i - lb + 1)
        End If
    Next i
    MovingAverage = result
End Function
' Binning — 等宽分箱
'
' 参数:
'   data  - 数值数据
'   nBins - 箱数
' 返回: Dictionary (binIndex 1..nBins → count)
'       附带键 "__internal_edges__" → 边界值数组
'=============================================================================
Public Function Binning( _
    ByRef data As Variant, _
    ByVal nBins As Long, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim result As Object
    Dim arrV As Variant
    Dim i As Long, lb As Long, ub As Long, n As Long
    Dim minVal As Double, maxVal As Double, binWidth As Double
    Dim binIdx As Long
    Dim edges1() As Double
    Dim ks As String
    Dim edges() As Double

    Set result = DP.Create()

    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 1 Or nBins < 1 Then
        Err.Raise ERR_INVALID_INPUT, "Binning", "数据为空或 nBins < 1。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    lb = LBound(arr): ub = UBound(arr)
    n = ub - lb + 1

    minVal = arr(lb): maxVal = arr(lb)
    For i = lb + 1 To ub
        If arr(i) < minVal Then minVal = arr(i)
        If arr(i) > maxVal Then maxVal = arr(i)
    Next i

    If minVal = maxVal Then
        For i = lb To ub
            binIdx = 1
            If result.Exists(CStr(binIdx)) Then
                result(CStr(binIdx)) = CLng(result(CStr(binIdx))) + 1
            Else
                result.Add CStr(binIdx), 1
            End If
        Next i
        ReDim edges1(0 To 1)
        edges1(0) = minVal - 0.5: edges1(1) = maxVal + 0.5
        result("__internal_edges__") = edges1
        Set Binning = result
        Exit Function
    End If

    binWidth = (maxVal - minVal) / CDbl(nBins)
    If binWidth <= 0# Then binWidth = 1#

    For i = lb To ub
        binIdx = Int((arr(i) - minVal) / binWidth) + 1
        If binIdx > nBins Then binIdx = nBins
        If binIdx < 1 Then binIdx = 1
        ks = CStr(binIdx)
        If result.Exists(ks) Then
            result(ks) = CLng(result(ks)) + 1
        Else
            result.Add ks, 1
        End If
    Next i

    ' 存储边界
    ReDim edges(0 To nBins)
    For i = 0 To nBins
        edges(i) = minVal + CDbl(i) * binWidth
    Next i
    result("__internal_edges__") = edges

    Set Binning = result
End Function

'=============================================================================


'=============================================================================
'===== 推断统计 (Inference) — ZTest → ConfidenceInterval =====
' ZTest — 单样本 Z 检验 (双侧 p 值，已知总体标准差)
'=============================================================================
Public Function ZTest(ByRef data As Variant, ByVal mu0 As Double, _
    ByVal sigma As Double, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant: arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 2 Then Err.Raise ERR_INVALID_INPUT, "ZTest", "需要至少 2 个有效数值。": Exit Function
    Dim arr() As Double: arr = arrV
    Dim n As Long: n = UBound(arr) - LBound(arr) + 1
    If sigma <= 0# Then Err.Raise ERR_INVALID_INPUT, "ZTest", "总体标准差 sigma 必须为正数。": Exit Function
    Dim m As Double: m = MeanDouble(arr)
    Dim z As Double: z = (m - mu0) / (sigma / Sqr(n))
    ZTest = 2# * (1# - NormSDistCDF(Abs(z)))
End Function

'=============================================================================
' TTest — 双样本 t 检验 (双侧 p 值，等方差假设)
' type: 1=paired, 2=equal variance, 3=unequal variance (Welch)
'=============================================================================
Public Function TTest(ByRef data1 As Variant, ByRef data2 As Variant, _
    Optional ByVal testType As Long = 2, Optional ByVal colIdx1 As Long = 1, Optional ByVal colIdx2 As Long = 1) As Variant
    Dim arrV1 As Variant, arrV2 As Variant: arrV1 = ExtractDoubles(data1, colIdx1): arrV2 = ExtractDoubles(data2, colIdx2)
    Dim n1 As Long: n1 = UBound(arrV1) - LBound(arrV1) + 1
    Dim n2 As Long: n2 = UBound(arrV2) - LBound(arrV2) + 1
    If n1 < 2 Or n2 < 2 Then Err.Raise ERR_INVALID_INPUT, "TTest", "每组需要至少 2 个有效数值。": Exit Function
    Dim a1() As Double, a2() As Double: a1 = arrV1: a2 = arrV2
    Dim m1 As Double, m2 As Double, s1 As Double, s2 As Double
    m1 = MeanDouble(a1): m2 = MeanDouble(a2): s1 = StdDevDouble(a1): s2 = StdDevDouble(a2)
    Dim t As Double, df As Double, se As Double
    Select Case testType
        Case 1: ' paired
            If n1 <> n2 Then Err.Raise ERR_INVALID_INPUT, "TTest", "配对检验需要两组样本量相同。": Exit Function
            Dim pdiff() As Double, i As Long: ReDim pdiff(LBound(a1) To UBound(a1))
            For i = LBound(a1) To UBound(a1): pdiff(i) = a1(i) - a2(i): Next i
            Dim md As Double: md = MeanDouble(pdiff): Dim sd As Double: sd = StdDevDouble(pdiff)
            If sd = 0# Then Err.Raise ERR_DIV_BY_ZERO, "TTest", "配对差值恒定，标准差为零。": Exit Function
            t = md / (sd / Sqr(n1)): df = n1 - 1
        Case 2: ' equal variance
            Dim sp2 As Double: sp2 = ((n1 - 1) * s1 ^ 2 + (n2 - 1) * s2 ^ 2) / (n1 + n2 - 2)
            se = Sqr(sp2 * (1# / n1 + 1# / n2))
            If se = 0# Then Err.Raise ERR_DIV_BY_ZERO, "TTest", "合并标准差为零，无法计算 t 统计量。": Exit Function
            t = (m1 - m2) / se: df = n1 + n2 - 2
        Case 3: ' Welch
            se = Sqr(s1 ^ 2 / n1 + s2 ^ 2 / n2)
            If se = 0# Then Err.Raise ERR_DIV_BY_ZERO, "TTest", "两组标准差均为零，无法计算 t 统计量。": Exit Function
            t = (m1 - m2) / se
            df = (s1 ^ 2 / n1 + s2 ^ 2 / n2) ^ 2 / _
                 ((s1 ^ 2 / n1) ^ 2 / (n1 - 1) + (s2 ^ 2 / n2) ^ 2 / (n2 - 1))
        Case Else: Err.Raise ERR_INVALID_INPUT, "TTest", "不支持的检验类型: " & testType: Exit Function
    End Select
    If df < 1 Then df = 1
    TTest = 2# * TDistCDF(Abs(t), df)  ' 保留 Double 精度, 不截断自由度 (#20)
End Function

'=============================================================================
' StandardError — 均值的标准误差 SE = s / sqrt(n)
'=============================================================================
Public Function StandardError(ByRef data As Variant, Optional ByVal colIndex As Long = 1) As Variant
    Dim arrV As Variant
    arrV = ExtractDoubles(data, colIndex)
    If UBound(arrV) - LBound(arrV) + 1 < 2 Then
        Err.Raise ERR_INVALID_INPUT, "StandardError", "需要至少 2 个有效数值。"
        Exit Function
    End If
    Dim arr() As Double
    arr = arrV
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1
    StandardError = StdDevDouble(arr) / Sqr(n)
End Function

'=============================================================================
' ConfidenceInterval — 置信区间 (t 分布近似)
'
' alpha = 0.05 → 95% 置信区间
' 返回字典: "lower", "upper", "mean", "se"
'=============================================================================
Public Function ConfidenceInterval( _
    ByRef data As Variant, _
    Optional ByVal alpha As Double = 0.05, _
    Optional ByVal colIndex As Long = 1) As Variant

    Dim result As Object
    Set result = DP.Create()
    Dim m As Double, se As Double, tVal As Double
    Dim n As Long

    Dim arrTmp As Variant
    arrTmp = ExtractDoubles(data, colIndex)
    n = UBound(arrTmp) - LBound(arrTmp) + 1

    If n < 2 Then
        Err.Raise ERR_INVALID_INPUT, "ConfidenceInterval", "需要至少 2 个数据点。"
        Exit Function
    End If
    If alpha <= 0# Or alpha >= 1# Then
        Err.Raise ERR_INVALID_INPUT, "ConfidenceInterval", "alpha 必须在 (0, 1) 之间。"
        Exit Function
    End If

    Dim arrTmpD() As Double
    arrTmpD = arrTmp
    m = MeanDouble(arrTmpD)
    se = StdDevDouble(arrTmpD) / Sqr(n)

    tVal = TDistCritical(alpha, n - 1)

    result("lower") = m - tVal * se
    result("upper") = m + tVal * se
    result("mean") = m
    result("se") = se
    Set ConfidenceInterval = result
End Function

'=============================================================================

'=============================================================================
'===== 分布函数 (Distribution) — NormSDistCDF =====
' NormSDistCDF — 标准正态分布累积概率（ZTest 调用）
'=============================================================================
Private Function NormSDistCDF(ByVal z As Double) As Double
    ' Abramowitz-Stegun 7.1.26 近似，误差 < 7.5e-8
    ' 使用 Horner 方法展开多项式，避免 VBA 表达式嵌套过深
    Dim t As Double, y As Double, p As Double
    t = 1# / (1# + 0.2316419 * Abs(z))
    y = 1.330274429#
    y = -1.821255978# + t * y
    y = 1.781477937# + t * y
    y = -0.356563782# + t * y
    y = 0.31938153# + t * y
    y = t * y
    p = y * Exp(-Abs(z) * Abs(z) / 2#) / Sqr(2# * 3.14159265358979)
    If z < 0# Then
        NormSDistCDF = p
    ElseIf p < TOL_STRICT Then
        NormSDistCDF = 1#  ' 避免 1-p 灾难性抵消 (p 在 Double 精度下已为零)
    Else
        NormSDistCDF = 1# - p
    End If
End Function

'=============================================================================
' GammaLn — 对数 Gamma 函数（Lanczos 近似）
'=============================================================================
Public Function GammaLn(ByVal x As Double) As Double
    ' 域校验: Gamma 函数在非正整数处有极点 (x = 0, -1, -2, ...)
    If x <= 0# Then
        If x = Int(x) Then
            Err.Raise ERR_INVALID_INPUT, "GammaLn", _
                "x 不得为非正整数 (Gamma 函数极点), 实际为 " & x
        End If
    End If
    ' Lanczos coefficients (g=7, n=9)
    Static coeffs As Variant
    Dim i As Long, ser As Double, tmp As Double
    If IsEmpty(coeffs) Then
        coeffs = Array(0.99999999999980993, 676.5203681218851, -1259.1392167224028, _
            771.32342877765313, -176.61502916214059, 12.507343278686905, _
            -0.13857109526572012, 9.9843695780195716E-06, 1.5056327351493116E-07)
    End If
    If x < 0.5 Then
        GammaLn = Log(3.14159265358979 / Sin(3.14159265358979 * x)) - GammaLn(1# - x)
        Exit Function
    End If
    x = x - 1#
    ser = coeffs(0)
    For i = 1 To 8
        ser = ser + coeffs(i) / (x + CDbl(i))
    Next i
    tmp = x + 7.5
    GammaLn = 0.5 * Log(2# * 3.14159265358979) + (x + 0.5) * Log(tmp) - tmp + Log(ser)
End Function

'=============================================================================
' TDistCritical — t 分布临界值 (双侧)（Private，ConfidenceInterval 调用）
'=============================================================================
Private Function TDistCritical(ByVal alpha As Double, ByVal df As Long) As Double
    If df < 1 Then df = 1
    If alpha <= 0# Or alpha >= 1# Then
        TDistCritical = 4#
        Exit Function
    End If
    TDistCritical = TInv2T(alpha, df)
End Function

'=============================================================================
' (NormSInv 已删除 — 无调用者的死代码; ConfidenceInterval 走 TDistCritical → TInv2T 路径)
'=============================================================================

'=============================================================================
' BetaReg — 正则化不完全 Beta 函数 I_x(a,b)（连分式展开）
'=============================================================================
'=============================================================================
' BetaReg — 正则化不完全 Beta 函数 I_x(a,b)
' 使用连分式展开 (Numerical Recipes §6.4 Modified Lentz method)
'=============================================================================
Private Function BetaCF(ByVal a As Double, ByVal b As Double, ByVal x As Double) As Double
    Const MAXIT As Long = 300
    ' Uses module-level NUM_EPS (machine epsilon) and FPMIN

    Dim qab As Double, qap As Double, qam As Double
    qab = a + b: qap = a + 1#: qam = a - 1#

    ' First step — evaluate 1/(1 - (a+b)/(a+1) * x)
    Dim c As Double, d As Double, f As Double
    c = 1#
    d = 1# - qab * x / qap
    If Abs(d) < FPMIN Then d = FPMIN
    d = 1# / d
    f = d

    Dim m As Long, m2 As Long
    Dim aa As Double, delta As Double, prevF As Double
    For m = 1 To MAXIT
        m2 = 2 * m

        ' Even step (d_{2m} with NR indexing)
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1# + aa * d
        If Abs(d) < FPMIN Then d = FPMIN
        c = 1# + aa / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1# / d
        f = f * (c * d)

        ' Odd step (d_{2m+1} with NR indexing)
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1# + aa * d
        If Abs(d) < FPMIN Then d = FPMIN
        c = 1# + aa / c
        If Abs(c) < FPMIN Then c = FPMIN
        d = 1# / d
        delta = c * d
        f = f * delta
        If Abs(delta - 1#) < NUM_EPS Then
            If m > 1 Then
                If Abs(f - prevF) < NUM_EPS * Abs(f) Then Exit For
            End If
            prevF = f
        End If
    Next m
    BetaCF = f
End Function

Public Function BetaReg(ByVal x As Double, ByVal a As Double, ByVal b As Double) As Double
    ' 域校验: a, b 必须为正 (GammaLn 在 0/负整数无定义)
    If a <= 0# Or b <= 0# Then
        Err.Raise ERR_INVALID_INPUT, "BetaReg", _
            "参数 a, b 必须为正数, 实际 a=" & a & ", b=" & b
    End If
    If x < 0# Or x > 1# Then BetaReg = -1#: Exit Function
    If x = 0# Then BetaReg = 0#: Exit Function
    If x = 1# Then BetaReg = 1#: Exit Function

    ' Symmetry: I_x(a,b) = 1 - I_{1-x}(b,a)
    If x > (a + 1#) / (a + b + 2#) Then
        BetaReg = 1# - BetaReg(1# - x, b, a)
        Exit Function
    End If

    ' bt = exp[lnΓ(a+b) - lnΓ(a) - lnΓ(b) + a*ln(x) + b*ln(1-x)]
    Dim bt As Double
    bt = Exp(GammaLn(a + b) - GammaLn(a) - GammaLn(b) _
           + a * Log(x) + b * Log(1# - x))
    If bt = 0# Then BetaReg = 0#: Exit Function

    BetaReg = bt * BetaCF(a, b, x) / a
    If BetaReg < 0# Then BetaReg = 0#
    If BetaReg > 1# Then BetaReg = 1#
End Function

'=============================================================================
' TDistCDF — t 分布上尾概率 P(T > |t|)（TTest / Binning / 外部调用）
'
' ⚠️ 命名说明: 虽然函数名含 "CDF"，但实际返回上尾概率（Survival Function）
'   即 P(T > |t|)，而非累积分布函数 P(T ≤ t)。
'   双尾检验应使用 TDist2T（= 2 * P(T > |t|)）。
'   参考: NormSDistCDF 返回的是真正的 CDF P(Z ≤ z)。
'=============================================================================
Public Function TDistCDF(ByVal tVal As Double, ByVal df As Double) As Double
    ' 纯 VBA 实现：P(T > |t|) = 0.5 * I_x(df/2, 0.5)  where x = df/(df + t²)
    ' df 接受 Double — Welch t-test 自由度可为非整数
    Dim x As Double: x = df / (df + tVal * tVal)
    Dim a As Double: a = df / 2#
    TDistCDF = 0.5 * BetaReg(x, a, 0.5)
End Function

'=============================================================================
' TDist2T — t 分布双尾 P 值 = 2 * P(T > |t|)
'=============================================================================
Public Function TDist2T(ByVal tStat As Double, ByVal df As Double) As Double
    If df <= 0 Then TDist2T = 1#: Exit Function
    TDist2T = 2# * TDistCDF(Abs(tStat), df)
End Function

'=============================================================================
' FDistRT — F 分布右尾概率 P(F > f)（通过不完全 Beta 函数）
'=============================================================================
Public Function FDistRT(ByVal fStat As Double, ByVal df1 As Double, ByVal df2 As Double) As Double
    If df1 <= 0 Or df2 <= 0 Then FDistRT = 1#: Exit Function
    If fStat <= 0# Then FDistRT = 1#: Exit Function
    ' P(F > f) = I_x(df2/2, df1/2) where x = df2 / (df2 + df1 * f)
    Dim x As Double: x = df2 / (df2 + df1 * fStat)
    FDistRT = BetaReg(x, df2 / 2#, df1 / 2#)
End Function

'=============================================================================
' TInv2T — t 分布双侧临界值（二分搜索，纯 VBA）
'=============================================================================
Public Function TInv2T(ByVal alpha As Double, ByVal df As Long) As Double
    If df < 1 Then df = 1
    If alpha <= 0# Or alpha >= 1# Then TInv2T = 4#: Exit Function
    ' 二分搜索求解 TDist2T(t, df) = alpha
    Dim lo As Double, hi As Double, mid As Double, p As Double
    Dim i As Long
    lo = 0#: hi = 100#  ' t=100 对所有 df 的 p 值 < 1E-200
    For i = 1 To 60
        mid = (lo + hi) / 2#
        p = TDist2T(mid, df)
        If p > alpha Then lo = mid Else hi = mid
        If Abs(hi - lo) < TOL_BISECT Then Exit For
    Next i
    TInv2T = (lo + hi) / 2#
End Function

'=============================================================================
'=============================================================================
' MaxAbsDouble — 数组最大绝对值 (尺度感知容差的参考尺度)
'=============================================================================
Private Function MaxAbsDouble(ByRef arr() As Double) As Double
    Dim i As Long, m As Double
    m = 0#
    For i = LBound(arr) To UBound(arr)
        If Abs(arr(i)) > m Then m = Abs(arr(i))
    Next i
    MaxAbsDouble = m
End Function

' QuickSortDouble — Hybrid QuickSort + InsertionSort for Double arrays (in-place, ascending)
' QuickSort for n>16, InsertionSort for n<=16. Recurses on smaller partition first → O(log n) stack.
'=============================================================================

Private Sub QuickSortDouble(ByRef arr() As Double, ByVal low As Long, ByVal high As Long)
    Dim lo As Long, hi As Long, pivot As Double, tmp As Double
    Dim i As Long, j As Long, key As Double

    If high - low <= DBL_INSERTION_THRESHOLD Then
        For i = low + 1 To high
            key = arr(i): j = i - 1
            Do While j >= low
                If arr(j) <= key Then Exit Do
                arr(j + 1) = arr(j): j = j - 1
            Loop
            arr(j + 1) = key
        Next i
        Exit Sub
    End If

    lo = low: hi = high
    ' 三数取中选轴 (SKILL §5.1): 避免中点轴在山形/模式化数据上退化 O(n²)
    Dim mid As Long: mid = (low + high) \ 2
    If arr(low) > arr(mid) Then tmp = arr(low): arr(low) = arr(mid): arr(mid) = tmp
    If arr(low) > arr(high) Then tmp = arr(low): arr(low) = arr(high): arr(high) = tmp
    If arr(mid) > arr(high) Then tmp = arr(mid): arr(mid) = arr(high): arr(high) = tmp
    pivot = arr(mid)
    Do While lo <= hi
        Do While arr(lo) < pivot: lo = lo + 1: Loop
        Do While arr(hi) > pivot: hi = hi - 1: Loop
        If lo <= hi Then
            tmp = arr(lo): arr(lo) = arr(hi): arr(hi) = tmp
            lo = lo + 1: hi = hi - 1
        End If
    Loop

    ' Recurse on smaller partition first
    If hi - low < high - lo Then
        If low < hi Then QuickSortDouble arr, low, hi
        If lo < high Then QuickSortDouble arr, lo, high
    Else
        If lo < high Then QuickSortDouble arr, lo, high
        If low < hi Then QuickSortDouble arr, low, hi
    End If
End Sub



'=============================================================================
' 工作表函数 (UDF_STAT_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_STAT_MEAN(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_MEAN = Mean(data, colIndex): Exit Function
EH: UDF_STAT_MEAN = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_MEDIAN(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_MEDIAN = Median(data, colIndex): Exit Function
EH: UDF_STAT_MEDIAN = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_STDEV(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_STDEV = StdDev(data, colIndex): Exit Function
EH: UDF_STAT_STDEV = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_STDEVP(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_STDEVP = StdDevP(data, colIndex): Exit Function
EH: UDF_STAT_STDEVP = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_VAR(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_VAR = Variance(data, colIndex): Exit Function
EH: UDF_STAT_VAR = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_VARP(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_VARP = VarianceP(data, colIndex): Exit Function
EH: UDF_STAT_VARP = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_PERCENTILE(ByVal data As Variant, ByVal k As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_PERCENTILE = Percentile(data, k, colIndex): Exit Function
EH: UDF_STAT_PERCENTILE = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_IQR(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_IQR = IQR(data, colIndex): Exit Function
EH: UDF_STAT_IQR = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_SKEW(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_SKEW = Skewness(data, colIndex): Exit Function
EH: UDF_STAT_SKEW = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_KURTOSIS(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_KURTOSIS = Kurtosis(data, colIndex): Exit Function
EH: UDF_STAT_KURTOSIS = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_MODE(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_MODE = Mode(data, colIndex): Exit Function
EH: UDF_STAT_MODE = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_MINMAX(ByVal data As Variant, _
    Optional ByVal outMin As Variant, _
    Optional ByVal outMax As Variant, _
    Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH
    Dim mn As Double, mx As Double
    UDF_STAT_MINMAX = MinMax(data, mn, mx, colIndex)
    Exit Function
EH: UDF_STAT_MINMAX = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_COV(ByVal dataX As Variant, ByVal dataY As Variant, _
    Optional ByVal colX As Variant = 1, Optional ByVal colY As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_COV = Covariance(dataX, dataY, colX, colY): Exit Function
EH: UDF_STAT_COV = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_CORREL(ByVal dataX As Variant, ByVal dataY As Variant, _
    Optional ByVal colX As Variant = 1, Optional ByVal colY As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_CORREL = Correlation(dataX, dataY, colX, colY): Exit Function
EH: UDF_STAT_CORREL = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_ZSCORE(ByVal data As Variant, _
    Optional ByVal value As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_ZSCORE = ZScore(data, value, colIndex): Exit Function
EH: UDF_STAT_ZSCORE = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_NORMALIZE(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_NORMALIZE = Normalize(data, colIndex): Exit Function
EH: UDF_STAT_NORMALIZE = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_LININTERP(ByVal x As Variant, ByVal xs As Variant, ByVal ys As Variant) As Variant
    On Error GoTo EH:     UDF_STAT_LININTERP = LinInterp(x, xs, ys): Exit Function
EH: UDF_STAT_LININTERP = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_GEOMEAN(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_GEOMEAN = GeometricMean(data, colIndex): Exit Function
EH: UDF_STAT_GEOMEAN = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_RANKEQ(ByVal data As Variant, ByVal value As Variant, _
    Optional ByVal ascending As Variant = True, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_RANKEQ = RankEq(data, value, ascending, colIndex): Exit Function
EH: UDF_STAT_RANKEQ = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_RANKAVG(ByVal data As Variant, ByVal value As Variant, _
    Optional ByVal ascending As Variant = True, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_RANKAVG = RankAvg(data, value, ascending, colIndex): Exit Function
EH: UDF_STAT_RANKAVG = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_HARMEAN(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_HARMEAN = HarmonicMean(data, colIndex): Exit Function
EH: UDF_STAT_HARMEAN = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_TRIMEAN(ByVal data As Variant, _
    Optional ByVal trimPct As Variant = 0.1, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_TRIMEAN = TrimMean(data, trimPct, colIndex): Exit Function
EH: UDF_STAT_TRIMEAN = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_RMS(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_RMS = RootMeanSquare(data, colIndex): Exit Function
EH: UDF_STAT_RMS = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_MAD(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_MAD = MeanAbsDev(data, colIndex): Exit Function
EH: UDF_STAT_MAD = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_SE(ByVal data As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_SE = StandardError(data, colIndex): Exit Function
EH: UDF_STAT_SE = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_R2(ByVal actual As Variant, ByVal predicted As Variant, _
    Optional ByVal colActual As Variant = 1, Optional ByVal colPredicted As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_R2 = RSquare(actual, predicted, colActual, colPredicted): Exit Function
EH: UDF_STAT_R2 = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_RANK(ByVal data As Variant, ByVal value As Variant, _
    Optional ByVal ascending As Variant = True, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_RANK = Rank(data, value, ascending, colIndex): Exit Function
EH: UDF_STAT_RANK = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_PERCENTRANK(ByVal data As Variant, ByVal value As Variant, _
    Optional ByVal ascending As Variant = True, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_PERCENTRANK = PercentRank(data, value, ascending, colIndex): Exit Function
EH: UDF_STAT_PERCENTRANK = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_WINSORIZE(ByVal data As Variant, _
    Optional ByVal pct As Variant = 0.05, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_WINSORIZE = Winsorize(data, pct, colIndex): Exit Function
EH: UDF_STAT_WINSORIZE = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_MA(ByVal data As Variant, ByVal window As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH:     UDF_STAT_MA = MovingAverage(data, window, colIndex): Exit Function
EH: UDF_STAT_MA = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_WMEAN(ByVal values As Variant, ByVal weights As Variant, _
    Optional ByVal colIdxV As Variant = 1, Optional ByVal colIdxW As Variant = 1) As Variant
    On Error GoTo EH: UDF_STAT_WMEAN = WeightedMean(values, weights, colIdxV, colIdxW): Exit Function
EH: UDF_STAT_WMEAN = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_ZTEST(ByVal data As Variant, ByVal mu0 As Variant, _
    ByVal sigma As Variant, Optional ByVal colIndex As Variant = 1) As Variant
    On Error GoTo EH: UDF_STAT_ZTEST = ZTest(data, mu0, sigma, colIndex): Exit Function
EH: UDF_STAT_ZTEST = CVErr(xlErrValue)
End Function

Public Function UDF_STAT_TTEST(ByVal data1 As Variant, ByVal data2 As Variant, _
    Optional ByVal testType As Variant = 2, Optional ByVal colIdx1 As Variant = 1, Optional ByVal colIdx2 As Variant = 1) As Variant
    On Error GoTo EH: UDF_STAT_TTEST = TTest(data1, data2, testType, colIdx1, colIdx2): Exit Function
EH: UDF_STAT_TTEST = CVErr(xlErrValue)
End Function

'=====================================================================
' 使用示例
'=====================================================================
' Mean(Range("A1:A100"))                                  → 均值
' GeometricMean(Array(1.05, 1.08, 1.03))                  → ~1.053
' Median(Array(1, 2, 3, 4, 100))                          → 3
' StdDev(Range("A1:A100"))                                 → 样本标准差
' IQR(Range("A1:A100"))                                    → 四分位距
' Skewness(Range("A1:A100"))                               → 偏度
' Correlation(Range("A1:A100"), Range("B1:B100"))          → 相关系数
' Rank(Range("A1:A10"), 50)                                → 排名
' LinInterp(2.5, Array(1,2,3), Array(10,20,30))            → 25
' Set bins = Binning(Range("A1:A100"), 10)                 → 等宽分箱字典

'=====================================================================