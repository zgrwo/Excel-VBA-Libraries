Option Explicit

'==============================================================================
' Module:       ArrayUtils
' Purpose:      Array operations: sort, filter, slice, aggregate, lookup
' Layer:        Data
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       62 functions/subs
'==============================================================================

Private DP As New DictProxy
Private AO As New ArrayOps

'=====================================================================
' ArrayUtils.bas — VBA 数组操作标准库
'
' 工作表函数 (UDF_ARR_*):
'   UDF_ARR_UNIQUE       — 去重    UDF_ARR_FILTER  — 条件筛选
'   UDF_ARR_SORT         — 排序    UDF_ARR_CONCAT  — 连接数组
'   UDF_ARR_SLICE        — 切片    UDF_ARR_FLATTEN — 展平
'   UDF_ARR_TRANSPOSE    — 1D→2D  UDF_ARR_REVERSE — 反转
'   UDF_ARR_LINSPACE     — 等间距  UDF_ARR_SAMPLE  — 抽样
'   UDF_ARR_FIND         — 查找    UDF_ARR_CONTAINS — 包含
'   UDF_ARR_SHUFFLE      — 洗牌    UDF_ARR_RANGEFILL — 序列
'   UDF_ARR_CHUNK        — 分块    UDF_ARR_LOOKUP  — 多列查找
'   UDF_ARR_MIN / MAX / SUM / PRODUCT / CUMSUM — 聚合
'   UDF_ARR_TOSTRING     — 转字符串
'   UDF_ARR_GETROW / GETCOL — 取行/列
'   UDF_ARR_TRANSPOSE2D  — 2D转置  UDF_ARR_EQUAL  — 比较
'   UDF_ARR_ARGSORT      — 排序索引
'   UDF_ARR_ANY / ALL    — 条件判断
'   UDF_ARR_COUNTIF      — 条件计数
'
' 约定:
'   - 所有输出数组均为 0-based（Transpose1D/2D 为兼容 VBA Range 使用 1-based）
'   - 所有函数为纯函数：返回新数组，绝不修改输入参数
'   - Null 排序优先于所有值；Error 排序次于所有值
'   - 未找到: ArrayFind 返回 -1；ArrayMin/Max 返回 Empty；其他返回空数组
'   - ArrayFind、ArraySlice、ArrayGetRow、ArrayGetCol、ArrayChunk 使用 0-based 索引
'   - Object 和嵌套数组元素安全处理（视为不透明值）
'   - 传入标量到筛选/检查函数时，按单元素数组处理
'   - Empty 在去重/排序中视为 0（因 IsNumeric(Empty) = True），但在聚合
'     (ArrayMin/Max/Sum/Product) 中通过 IsEmpty 显式排除，作为缺失数据处理
'
' 错误码:
'   vbObjectError + 1001 — 需要 1D 数组
'   vbObjectError + 1002 — 需要 2D 数组
'   vbObjectError + 1003 — 索引/参数越界
'   vbObjectError + 1004 — 无效输入 / 外部组件不可用
'
' 一维操作:
'   ArrayUnique       — 去重（保留首次出现顺序）
'   ArraySort         — QuickSort 排序（O(n log n)，不保证稳定排序）
'   ArrayFilterByValue — 按值与比较运算符筛选
'   ArrayCountIf      — 统计符合条件的元素数量
'   ArrayConcat       — 连接两个一维数组
'   ArraySlice        — 提取子数组（Python 风格索引）
'   ArrayReverse      — 反转顺序
'
' 二维操作:
'   ArrayFlatten      — 二维展平为一维（行优先）
'   ArrayTranspose1D  — 一维转二维列/行向量（1-based 输出）
'   ArrayTranspose2D  — 二维行列转置（1-based 输出）
'   ArrayGetRow       — 提取单行（0-based 索引）
'   ArrayGetCol       — 提取单列（0-based 索引）
'   ArrayChunk        — 按固定大小分块
'
' 搜索 / 比较:
'   ArrayFind         — 查找首个索引（0-based 偏移，未找到返回 -1）
'   ArrayContains     — 判断是否包含
'   ArrayLookup       — 内存数组版多列查找（VLOOKUP 替代）
'   ArrayEqual        — 逐元素相等比较
'   ArrayAny          — 是否存在任意元素满足条件
'   ArrayAll          — 是否所有元素都满足条件
'
' 数值:
'   ArrayMin / ArrayMax / ArraySum — 数值聚合
'   ArrayProduct      — 数值元素连乘积
'   CumSum            — 累积和
'   ArgSort           — 返回排序后的索引
'   LinSpace          — 生成 n 个等间距点
'
' 工具:
'   ArrayDims         — 数组维数（未初始化/标量返回 0）
'   RangeFill         — 等差数列生成器
'   ArrayShuffle      — Fisher-Yates 随机打乱（每次会话初始化一次随机种子）
'   ArraySample       — 随机抽样（支持有/无放回）
'   ArrayToString     — 用分隔符连接为字符串
'   IsArray1D         — 判断是否为一维数组
'=====================================================================

' -- 错误代码常量 -----------------------------------------------------
Private Const ERR_1D_REQUIRED   As Long = vbObjectError + 1001
Private Const ERR_2D_REQUIRED   As Long = vbObjectError + 1002
Private Const ERR_OUT_OF_BOUNDS As Long = vbObjectError + 1003
Private Const ERR_INVALID_INPUT As Long = vbObjectError + 1004
Private Const ERR_MULTI_AREA    As Long = vbObjectError + 1005

'=============================================================================
' NormalizeToArray — Range 则提取 .Value 为数组, 单列/单行则转一维
' NOTE: This duplicates VariantKit.Normalize1D/NormalizeTo2D functionality.
'   Delegation deferred — callers depend on exact return-type semantics.
'   Future: replace body with `VK.Normalize1D(v, "R")` after crossval regression.
'=============================================================================
Private Function NormalizeToArray(ByRef v As Variant) As Variant
    Dim rv As Variant, tmp() As Variant
    Dim nRows As Long, nCols As Long, i As Long

    If IsObject(v) Then
        If TypeOf v Is Range Then
            ' 多区域 Range（Ctrl+选择）会让 .Value 抛出错误 1004
            If v.Areas.Count > 1 Then
                NormalizeToArray = CVErr(xlErrValue)
                Exit Function
            End If
            rv = v.Value
            If IsArray(rv) Then
                nRows = UBound(rv, 1) - LBound(rv, 1) + 1
                nCols = UBound(rv, 2) - LBound(rv, 2) + 1
                If nRows = 1 And nCols = 1 Then
                    NormalizeToArray = rv(LBound(rv, 1), LBound(rv, 2))
                ElseIf nRows = 1 Then
                    ReDim tmp(LBound(rv, 2) To UBound(rv, 2))
                    For i = LBound(rv, 2) To UBound(rv, 2)
                        tmp(i) = rv(LBound(rv, 1), i)
                    Next i
                    NormalizeToArray = tmp
                ElseIf nCols = 1 Then
                    ReDim tmp(LBound(rv, 1) To UBound(rv, 1))
                    For i = LBound(rv, 1) To UBound(rv, 1)
                        tmp(i) = rv(i, LBound(rv, 2))
                    Next i
                    NormalizeToArray = tmp
                Else
                    NormalizeToArray = rv
                End If
            Else
                NormalizeToArray = rv
            End If
        Else
            NormalizeToArray = v
        End If
    Else
        NormalizeToArray = v
    End If
End Function

'=============================================================================
' ArrayDims — 获取数组维数 (未初始化数组返回 0，标量返回 0)
'=============================================================================
Public Function ArrayDims(ByRef arr As Variant) As Long
    Dim d As Long, ub As Long

    If Not IsArray(arr) Then
        ArrayDims = 0
        Exit Function
    End If

    ' 检测未初始化数组 (Dim arr() As Variant)
    ' LBound 对已分配空数组成功 (arr = Array())，对未初始化数组失败
    Err.Clear
    On Error Resume Next
    d = LBound(arr, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ArrayDims = 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    d = 2
    Do
        On Error Resume Next
        ub = UBound(arr, d)
        If Err.Number <> 0 Then Exit Do
        On Error GoTo 0
        d = d + 1
    Loop
    Err.Clear
    On Error GoTo 0
    ArrayDims = d - 1
End Function

'=============================================================================
' IsArray1D — 判断是否为 1D 数组 (未初始化数组返回 False)
'=============================================================================
Public Function IsArray1D(ByRef arr As Variant) As Boolean
    If Not IsArray(arr) Then
        IsArray1D = False
        Exit Function
    End If

    ' 检测未初始化数组: LBound 对已分配空数组成功，对未初始化数组失败
    Err.Clear
    On Error Resume Next
    Dim lb1 As Long
    lb1 = LBound(arr, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsArray1D = False
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    Dim ub2 As Long
    On Error Resume Next
    ub2 = UBound(arr, 2)
    If Err.Number <> 0 Then
        IsArray1D = True
    Else
        IsArray1D = False
    End If
    Err.Clear
    On Error GoTo 0
End Function

'=============================================================================
' 内部辅助函数
'=============================================================================

' SafeKey — 构建用于去重字典的稳定字符串 key
'
' 设计说明: SafeKey 在 IsNumeric 分支下将 Boolean(True→-1) 和 Empty(→0)
'   视为数值，因此 True 与 -1、Empty 与 0 在去重中合并。这是有意为之的
'   设计选择（非缺陷），因为 VBA 本身允许 Boolean/Empty 参与算术运算。
'   对比: ValuesEqual 对 Boolean/Empty 的检查更为严格——它要求类型完全匹配。
'   两种语义服务于不同场景：去重偏宽容合并，逐元素相等比较偏精确匹配。
'
' SafeKey — 委托给 VariantKit.SafeKey (单一数据源)
Private Function SafeKey(ByVal v As Variant) As String
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    SafeKey = vk.SafeKey(v)
End Function

Private Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Then
        SafeStr = "Null"
    ElseIf IsError(v) Then
        SafeStr = "#Error"
    ElseIf IsObject(v) Then
        SafeStr = "[Object]"
    ElseIf IsArray(v) Then
        SafeStr = "[Array]"
    Else
        SafeStr = CStr(v)
    End If
End Function

'=============================================================================
'=============================================================================
' ValuesEqual — 委托给 VariantKit.ValuesEqual (单一数据源, 含防御性 CDate 守卫)
'=============================================================================
Private Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    ValuesEqual = vk.ValuesEqual(a, b)
End Function

'=============================================================================

'=============================================================================
' WrapScalar — 将非数组标量包装为单元素数组，方便统一处理
' WrapScalar — 委托给 VariantKit.WrapScalar (单一数据源)
'=============================================================================
Private Function WrapScalar(ByRef v As Variant) As Variant
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    WrapScalar = vk.WrapScalar(v)
End Function

'=============================================================================
' CompareValues — 比较两个 Variant（返回 -1 / 0 / 1）
'
' 排序规则: Null < 数值/字符串/日期 < Error
' 日期比较仅当两个值都为 vbDate 类型时触发
'=============================================================================
Private Function CompareValues(ByVal a As Variant, ByVal b As Variant) As Long
    If IsNull(a) And IsNull(b) Then CompareValues = 0: Exit Function
    If IsNull(a) Then CompareValues = -1: Exit Function
    If IsNull(b) Then CompareValues = 1: Exit Function
    If IsError(a) And IsError(b) Then CompareValues = 0: Exit Function
    If IsError(a) Then CompareValues = 1: Exit Function
    If IsError(b) Then CompareValues = -1: Exit Function
    ' Empty: type-discriminated to prevent IsNumeric(Empty)=True from merging with 0
    If IsEmpty(a) And IsEmpty(b) Then CompareValues = 0: Exit Function
    If IsEmpty(a) Then CompareValues = -1: Exit Function
    If IsEmpty(b) Then CompareValues = 1: Exit Function
    ' Boolean: type-discriminated to prevent IsNumeric(True)=True from merging with -1
    If VarType(a) = vbBoolean And VarType(b) = vbBoolean Then
        If CBool(a) = CBool(b) Then
            CompareValues = 0
        ElseIf CBool(a) Then
            CompareValues = 1   ' True > False
        Else
            CompareValues = -1  ' False < True
        End If
        Exit Function
    End If
    If VarType(a) = vbBoolean Then CompareValues = 1: Exit Function
    If VarType(b) = vbBoolean Then CompareValues = -1: Exit Function
    If VarType(a) = vbDate And VarType(b) = vbDate Then
        If CDate(a) < CDate(b) Then
            CompareValues = -1
        ElseIf CDate(a) > CDate(b) Then
            CompareValues = 1
        Else
            CompareValues = 0
        End If
    ElseIf IsNumeric(a) And IsNumeric(b) Then
        If CDbl(a) < CDbl(b) Then
            CompareValues = -1
        ElseIf CDbl(a) > CDbl(b) Then
            CompareValues = 1
        Else
            CompareValues = 0
        End If
    Else
        CompareValues = StrComp(SafeStr(a), SafeStr(b), vbTextCompare)
    End If
End Function

'=============================================================================
' CompareAtIndices — 通过索引比较数组中的两个元素 (委托给 CompareValues)
'=============================================================================
Private Function CompareAtIndices(ByRef arr As Variant, ByVal idxA As Long, ByVal idxB As Long) As Long
    CompareAtIndices = CompareValues(arr(idxA), arr(idxB))
End Function

'=============================================================================
' FilterPasses / FilterEquals — 筛选条件判断
'=============================================================================
' FilterPasses — delegates to VariantKit.FilterPasses (single source of truth)
Private Function FilterPasses(ByVal element As Variant, ByVal matchValue As Variant, ByVal operator As String) As Boolean
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    FilterPasses = vk.FilterPasses(element, matchValue, operator)
End Function

Private Function FilterEquals(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) And IsNull(b) Then FilterEquals = True: Exit Function
    If IsNull(a) Or IsNull(b) Then FilterEquals = False: Exit Function
    If IsError(a) And IsError(b) Then FilterEquals = True: Exit Function
    If IsError(a) Or IsError(b) Then FilterEquals = False: Exit Function
    If IsNumeric(a) And IsNumeric(b) Then
        FilterEquals = (Abs(CDbl(a) - CDbl(b)) < 1E-12)
    Else
        FilterEquals = (StrComp(SafeStr(a), SafeStr(b), vbTextCompare) = 0)
    End If
End Function

'=============================================================================
' CheckIs1D — 检查输入是否为一维数组 (False=2D+, 按 SKILL.md 约定返回 CVErr)
'=============================================================================
Private Function CheckIs1D(ByRef arr As Variant) As Boolean
    ' Returns True if arr is 1D or not an array; False if 2D+
    If IsArray(arr) Then
        CheckIs1D = IsArray1D(arr)
    Else
        CheckIs1D = True
    End If
End Function

'=============================================================================
' ArrayUnique — 一维数组去重 (保留首次出现顺序, 输出 0-based)
'=============================================================================
Public Function ArrayUnique(ByRef arr As Variant) As Variant
    Dim dict As Object
    Dim i As Long, lb As Long, ub As Long
    Dim key As String
    Dim result() As Variant
    Dim v As Variant, idx As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayUnique = data
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayUnique", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    Set dict = DP.Create()

    For i = lb To ub
        key = SafeKey(data(i))
        If Not dict.Exists(key) Then
            dict.Add key, data(i)
        End If
    Next i

    If dict.Count = 0 Then
        result = Array()
        ArrayUnique = result
        Set dict = Nothing
        Exit Function
    End If

    ReDim result(0 To dict.Count - 1)
    idx = 0
    For Each v In dict.Items
        result(idx) = v
        idx = idx + 1
    Next v
    Set dict = Nothing
    ArrayUnique = result
End Function

'=============================================================================
' ArraySort — 一维数组排序（QuickSort, O(n log n), 输出 0-based）
'
' 支持: 数值、字符串、日期。
' Null 值排在最前，Error 值排在最后。
'=============================================================================
Public Function ArraySort( _
    ByRef arr As Variant, _
    Optional ByVal ascending As Boolean = True) As Variant

    Dim data As Variant
    Dim result() As Variant
    Dim lb As Long, ub As Long
    Dim i As Long

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArraySort = data
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArraySort", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    If ub < lb Then
        ' 空数组 — 按 SKILL.md 约定返回空 Variant 数组
        ArraySort = Array()
        Exit Function
    End If

    ReDim result(0 To ub - lb)
    For i = lb To ub
        If IsObject(data(i)) Then Set result(i - lb) = data(i) Else result(i - lb) = data(i)
    Next i

    If ub > lb Then
        AO.Sort result, ascending
    End If
    ArraySort = result
End Function

'=============================================================================
' ArrayFilterByValue — 按值筛选（返回符合条件的元素, 输出 0-based）
'
' 参数:
'   arr       - 一维数组
'   matchValue - 匹配值
'   operator  - "="（默认）, "<", ">", "<=", ">=", "<>", "contains", "regex"
'=============================================================================
Public Function ArrayFilterByValue( _
    ByRef arr As Variant, _
    ByVal matchValue As Variant, _
    Optional ByVal operator As String = "=") As Variant

    Dim result() As Variant
    Dim lb As Long, ub As Long
    Dim i As Long, cnt As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ' 标量按单元素数组处理 — 应用筛选条件后返回
        If FilterPasses(data, matchValue, operator) Then
            ReDim result(0 To 0)
            result(0) = data
            ArrayFilterByValue = result
        Else
            ArrayFilterByValue = Array()
        End If
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayFilterByValue", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)

    For i = lb To ub
        If FilterPasses(data(i), matchValue, operator) Then
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        result = Array()
        ArrayFilterByValue = result
        Exit Function
    End If

    ReDim result(0 To cnt - 1)
    cnt = 0
    For i = lb To ub
        If FilterPasses(data(i), matchValue, operator) Then
            result(cnt) = data(i)
            cnt = cnt + 1
        End If
    Next i
    ArrayFilterByValue = result
End Function

'=============================================================================
' ArrayCountIf — 统计符合条件的元素数量
'
' 参数与 ArrayFilterByValue 相同，但返回计数而非筛选结果。
'=============================================================================
Public Function ArrayCountIf( _
    ByRef arr As Variant, _
    ByVal matchValue As Variant, _
    Optional ByVal operator As String = "=") As Long
    Dim data As Variant, lb As Long, ub As Long, i As Long, cnt As Long
    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        If FilterPasses(data, matchValue, operator) Then ArrayCountIf = 1 Else ArrayCountIf = 0
        Exit Function
    End If
    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayCountIf", "需要一维数组。"
    lb = LBound(data): ub = UBound(data)
    For i = lb To ub
        If FilterPasses(data(i), matchValue, operator) Then cnt = cnt + 1
    Next i
    ArrayCountIf = cnt
End Function

'=============================================================================
' ArrayConcat — 连接两个一维数组 (输出始终 0-based)
'=============================================================================
Public Function ArrayConcat(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant
    Dim result() As Variant
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    Dim i As Long, idx As Long, n1 As Long, n2 As Long
    Dim data1 As Variant, data2 As Variant

    ' Normalize Range→Array, then wrap scalars for unified processing
    data1 = NormalizeToArray(arr1)
    data2 = NormalizeToArray(arr2)
    data1 = WrapScalar(data1)
    data2 = WrapScalar(data2)

    ' Compute bounds early — empty arrays (UBound < LBound) cannot be indexed
    lb1 = LBound(data1): ub1 = UBound(data1)
    lb2 = LBound(data2): ub2 = UBound(data2)
    n1 = ub1 - lb1 + 1: n2 = ub2 - lb2 + 1
    If n1 + n2 = 0 Then
        ArrayConcat = Array()
        Exit Function
    End If

    ' 传播 NormalizeToArray 的多区域 Range 错误 (WrapScalar 会将 CVErr 包装为数组)
    If n1 > 0 Then
        If IsError(data1(lb1)) Then ArrayConcat = data1(lb1): Exit Function
    End If
    If n2 > 0 Then
        If IsError(data2(lb2)) Then ArrayConcat = data2(lb2): Exit Function
    End If

    If Not CheckIs1D(data1) Then Err.Raise ERR_1D_REQUIRED, "ArrayConcat", "需要一维数组。"
    If Not CheckIs1D(data2) Then Err.Raise ERR_1D_REQUIRED, "ArrayConcat", "需要一维数组。"

    ReDim result(0 To n1 + n2 - 1)
    idx = 0
    For i = lb1 To ub1: result(idx) = data1(i): idx = idx + 1: Next i
    For i = lb2 To ub2: result(idx) = data2(i): idx = idx + 1: Next i
    ArrayConcat = result
End Function

'=============================================================================
' ArraySlice — 取一维子数组（输出 0-based）
'
' 参数:
'   start - 起始索引（0-based, 默认 0 即第一个元素; 负数从末尾倒数）
'   cnt   - 元素数量（默认 -1 表示取到末尾）
'=============================================================================
Public Function ArraySlice( _
    ByRef arr As Variant, _
    Optional ByVal start As Long = 0, _
    Optional ByVal cnt As Long = -1) As Variant

    Dim result() As Variant
    Dim lb As Long, ub As Long
    Dim i As Long, actualStart As Long, actualEnd As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArraySlice = data
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArraySlice", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)

    ' 空数组检测（lb > ub）
    If ub < lb Then
        result = Array()
        ArraySlice = result
        Exit Function
    End If

    If start >= 0 Then
        actualStart = lb + start
    Else
        actualStart = ub + start + 1
    End If

    If actualStart < lb Then actualStart = lb
    If actualStart > ub Then actualStart = ub

    If cnt < 0 Then
        actualEnd = ub
    Else
        actualEnd = actualStart + cnt - 1
        If actualEnd > ub Then actualEnd = ub
    End If

    If actualEnd < actualStart Then
        result = Array()
        ArraySlice = result
        Exit Function
    End If
    ReDim result(0 To actualEnd - actualStart)
    Dim idx As Long: idx = 0
    For i = actualStart To actualEnd
        result(idx) = data(i)
        idx = idx + 1
    Next i
    ArraySlice = result
End Function

'=============================================================================
' ArrayFlatten — 二维数组按行展平为一维 (输出 0-based)
'=============================================================================
Public Function ArrayFlatten(ByRef arr As Variant) As Variant
    Dim result() As Variant
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim r0 As Long, c0 As Long
    Dim idx As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayFlatten = data
        Exit Function
    End If

    If IsArray1D(data) Then
        ArrayFlatten = data
        Exit Function
    End If

    If ArrayDims(data) <> 2 Then Err.Raise ERR_2D_REQUIRED, "ArrayFlatten", "需要二维数组。": Exit Function

    r0 = LBound(data, 1): c0 = LBound(data, 2)
    nRows = UBound(data, 1) - r0 + 1
    nCols = UBound(data, 2) - c0 + 1

    ReDim result(0 To nRows * nCols - 1)
    idx = 0
    For i = 1 To nRows
        For j = 1 To nCols
            result(idx) = data(r0 + i - 1, c0 + j - 1)
            idx = idx + 1
        Next j
    Next i
    ArrayFlatten = result
End Function

'=============================================================================
' ArrayTranspose1D — 一维数组行列互转（输出 1-based, 遵循 VBA Range 惯例）
'
' 输入: 一维数组（如 Split 返回的 0-based 数组）
' 输出: 二维数组 (1 To n, 1 To 1) — 即列向量
'       或 (1 To 1, 1 To n) — 行向量
'=============================================================================
Public Function ArrayTranspose1D( _
    ByRef arr As Variant, _
    Optional ByVal asColumn As Boolean = True) As Variant

    Dim result() As Variant
    Dim lb As Long, ub As Long, n As Long
    Dim i As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayTranspose1D = data
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayTranspose1D", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    n = ub - lb + 1
    If n = 0 Then
        ArrayTranspose1D = Array()
        Exit Function
    End If

    If asColumn Then
        ReDim result(1 To n, 1 To 1)
        For i = 1 To n: result(i, 1) = data(lb + i - 1): Next i
    Else
        ReDim result(1 To 1, 1 To n)
        For i = 1 To n: result(1, i) = data(lb + i - 1): Next i
    End If
    ArrayTranspose1D = result
End Function

'=============================================================================
' ArrayFind — 查找元素索引 (未找到返回 -1)
'=============================================================================
Public Function ArrayFind( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    Optional ByVal caseSensitive As Boolean = False) As Variant

    Dim lb As Long, ub As Long, i As Long
    Dim compare As VbCompareMethod
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayFind = -1
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayFind", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    If caseSensitive Then compare = vbBinaryCompare Else compare = vbTextCompare

    For i = lb To ub
        If IsNull(data(i)) And IsNull(value) Then
            ArrayFind = i - lb
            Exit Function
        ElseIf IsNull(data(i)) Or IsNull(value) Then
            ' 跳过 — Null 只等于 Null
        ElseIf IsError(data(i)) And IsError(value) Then
            ' 精确匹配错误码 — 与 VariantKit.ValuesEqual 语义一致
            Dim errA As Long, errB As Long
            On Error Resume Next
            errA = data(i): errB = value
            On Error GoTo 0
            If errA = errB Then
                ArrayFind = i - lb
                Exit Function
            End If
        ElseIf IsError(data(i)) Or IsError(value) Then
            ' 跳过 — Error 只等于同样的 Error
        ElseIf IsEmpty(data(i)) And IsEmpty(value) Then
            ArrayFind = i - lb
            Exit Function
        ElseIf IsEmpty(data(i)) Or IsEmpty(value) Then
            ' 跳过 — Empty 仅匹配 Empty，不匹配 0
        ElseIf VarType(data(i)) <> vbBoolean And VarType(value) <> vbBoolean And IsNumeric(data(i)) And IsNumeric(value) Then
            If Abs(CDbl(data(i)) - CDbl(value)) < 1E-12 Then
                ArrayFind = i - lb
                Exit Function
            End If
        Else
            If StrComp(SafeStr(data(i)), SafeStr(value), compare) = 0 Then
                ArrayFind = i - lb
                Exit Function
            End If
        End If
    Next i
    ArrayFind = -1
End Function

'=============================================================================
' ArrayContains — 判断是否包含
'=============================================================================
Public Function ArrayContains( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    Optional ByVal caseSensitive As Boolean = False) As Variant

    Dim findResult As Variant
    findResult = ArrayFind(arr, value, caseSensitive)
    If IsError(findResult) Then
        ArrayContains = findResult
    Else
        ArrayContains = (findResult <> -1)
    End If
End Function

'=============================================================================
' ArrayLookup — 内存数组版多列查找
'
' 在 lookupArray 中查找 lookupValue，返回 returnCols 指定的列值。
' 参数: lookupArray(2D), lookupValue, lookupCol(1-based),
'        returnCols(列号或数组), matchType(0=精确/1=近似)
' 返回: 单值 Variant 或 2D 行数组 (1 To 1, 1 To nCols)
'=============================================================================
Public Function ArrayLookup( _
    ByRef lookupArray As Variant, _
    ByVal lookupValue As Variant, _
    ByVal lookupCol As Long, _
    Optional ByVal returnCols As Variant, _
    Optional ByVal matchType As Long = 0) As Variant

    Dim nRows As Long, nCols As Long, i As Long, j As Long
    Dim foundRow As Long, nRetCols As Long
    Dim retColArr() As Long, result As Variant
    Dim data As Variant
    Dim lb1 As Long, lb2 As Long   ' lower bounds (handle 0-based COM arrays)

    ' Extract Range .Value, preserving 2D structure
    If IsObject(lookupArray) Then
        If TypeOf lookupArray Is Range Then
            If lookupArray.Areas.Count > 1 Then Err.Raise ERR_MULTI_AREA, "ArrayLookup", "不支持多重区域。": Exit Function
            If lookupArray.Count = 1 Then
                ReDim data(1 To 1, 1 To 1): data(1, 1) = lookupArray.Value
            Else
                data = lookupArray.Value
            End If
        Else
            Err.Raise ERR_INVALID_INPUT, "ArrayLookup", "不支持的 lookupArray 类型。"
        End If
    Else
        data = lookupArray
    End If

    If Not IsArray(data) Then Err.Raise ERR_INVALID_INPUT, "ArrayLookup", "需要数组输入。"
    Err.Clear: On Error Resume Next
    lb2 = LBound(data, 2): nCols = UBound(data, 2) - lb2 + 1
    If Err.Number <> 0 Then Err.Raise ERR_2D_REQUIRED, "ArrayLookup", "需要二维数组。"
    On Error GoTo 0
    lb1 = LBound(data, 1): nRows = UBound(data, 1) - lb1 + 1
    If lookupCol < 1 Or lookupCol > nCols Then Err.Raise ERR_OUT_OF_BOUNDS, "ArrayLookup", "查找列越界: " & lookupCol

    ' Handle Range for returnCols (COM callers may pass Range objects)
    If IsObject(returnCols) Then
        If TypeOf returnCols Is Range Then returnCols = returnCols.Value
    End If

    ' 解析返回列
    If IsMissing(returnCols) Then
        nRetCols = nCols: ReDim retColArr(1 To nCols)
        For j = 1 To nCols: retColArr(j) = j: Next j
    ElseIf IsArray(returnCols) Then
        ' Flatten 1-row Range.Value (2D array) to 1D for correct UBound counting
        If ArrayDims(returnCols) = 2 Then
            Dim rcR As Long, rcC As Long, rcIdx As Long
            rcR = UBound(returnCols, 1) - LBound(returnCols, 1) + 1
            rcC = UBound(returnCols, 2) - LBound(returnCols, 2) + 1
            If rcR = 1 Then
                ' Row vector: flatten columns to 1D
                Dim tmpRC() As Variant
                ReDim tmpRC(LBound(returnCols, 2) To UBound(returnCols, 2))
                For rcIdx = LBound(returnCols, 2) To UBound(returnCols, 2)
                    tmpRC(rcIdx) = returnCols(LBound(returnCols, 1), rcIdx)
                Next rcIdx
                returnCols = tmpRC
            ElseIf rcC = 1 Then
                ' Column vector: flatten rows to 1D
                ReDim tmpRC(LBound(returnCols, 1) To UBound(returnCols, 1))
                For rcIdx = LBound(returnCols, 1) To UBound(returnCols, 1)
                    tmpRC(rcIdx) = returnCols(rcIdx, LBound(returnCols, 2))
                Next rcIdx
                returnCols = tmpRC
            End If
        End If
        nRetCols = UBound(returnCols) - LBound(returnCols) + 1
        ReDim retColArr(1 To nRetCols)
        For j = 1 To nRetCols: retColArr(j) = CLng(returnCols(LBound(returnCols) + j - 1)): Next j
    Else
        nRetCols = 1: ReDim retColArr(1 To 1): retColArr(1) = CLng(returnCols)
    End If
    For j = 1 To nRetCols
        If retColArr(j) < 1 Or retColArr(j) > nCols Then Err.Raise ERR_OUT_OF_BOUNDS, "ArrayLookup", "返回列越界: " & retColArr(j)
    Next j

    ' 查找
    foundRow = -1
    If matchType = 0 Then
        For i = 1 To nRows
            If ValuesEqual(data(lb1 + i - 1, lb2 + lookupCol - 1), lookupValue) Then foundRow = i: Exit For
        Next i
    Else
        For i = 1 To nRows
            If VarType(data(lb1 + i - 1, lb2 + lookupCol - 1)) <> vbBoolean And VarType(lookupValue) <> vbBoolean And _
               IsNumeric(data(lb1 + i - 1, lb2 + lookupCol - 1)) And IsNumeric(lookupValue) Then
                If CDbl(data(lb1 + i - 1, lb2 + lookupCol - 1)) <= CDbl(lookupValue) Then foundRow = i Else Exit For
            End If
        Next i
    End If
    If foundRow = -1 Then ArrayLookup = CVErr(xlErrNA): Exit Function

    If nRetCols = 1 Then
        ArrayLookup = data(lb1 + foundRow - 1, lb2 + retColArr(1) - 1)
    Else
        ReDim result(1 To 1, 1 To nRetCols)
        For j = 1 To nRetCols: result(1, j) = data(lb1 + foundRow - 1, lb2 + retColArr(j) - 1): Next j
        ArrayLookup = result
    End If
End Function

'=============================================================================
' ArrayShuffle — Fisher-Yates 随机打乱 (不修改原数组, 输出 0-based)
'=============================================================================
Public Function ArrayShuffle(ByRef arr As Variant) As Variant
    Dim result() As Variant
    Dim lb As Long, ub As Long, n As Long
    Dim i As Long, j As Long
    Dim tmp As Variant
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayShuffle = data
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayShuffle", "需要一维数组。"

    Static seeded As Boolean
    If Not seeded Then
        Randomize
        seeded = True
    End If

    lb = LBound(data): ub = UBound(data)
    If ub < lb Then
        ArrayShuffle = Array()
        Exit Function
    End If
    n = ub - lb
    ReDim result(0 To n)
    For i = lb To ub
        If IsObject(data(i)) Then Set result(i - lb) = data(i) Else result(i - lb) = data(i)
    Next i

    For i = n To 1 Step -1
        j = Int(Rnd() * (i + 1))
        If IsObject(result(i)) Then
            Set tmp = result(i): Set result(i) = result(j): Set result(j) = tmp
        Else
            tmp = result(i): result(i) = result(j): result(j) = tmp
        End If
    Next i
    ArrayShuffle = result
End Function

'=============================================================================
' ArraySample — 从数组中随机抽样 (无放回/有放回)
'=============================================================================
Public Function ArraySample(ByRef arr As Variant, _
    ByVal n As Long, _
    Optional ByVal withReplacement As Boolean = False) As Variant
    Dim data As Variant, result() As Variant, lb As Long, ub As Long, size As Long
    Dim i As Long, j As Long, tmp As Variant, idx As Long, rIdx As Long
    Static seeded As Boolean: If Not seeded Then Randomize: seeded = True
    data = NormalizeToArray(arr)
    If Not IsArray(data) Or n <= 0 Then ArraySample = Array(): Exit Function
    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArraySample", "需要一维数组。"
    lb = LBound(data): ub = UBound(data): size = ub - lb + 1
    If Not withReplacement And n > size Then n = size
    ReDim result(0 To n - 1)
    If withReplacement Then
        For i = 0 To n - 1
            rIdx = lb + Int(Rnd() * size)
            If IsObject(data(rIdx)) Then Set result(i) = data(rIdx) Else result(i) = data(rIdx)
        Next i
    Else
        Dim pool() As Variant: ReDim pool(0 To size - 1)
        For i = 0 To size - 1
            If IsObject(data(lb + i)) Then Set pool(i) = data(lb + i) Else pool(i) = data(lb + i)
        Next i
        For i = 0 To n - 1
            j = Int(Rnd() * (size - i))
            If IsObject(pool(j)) Then Set result(i) = pool(j) Else result(i) = pool(j)
            If IsObject(pool(size - i - 1)) Then Set pool(j) = pool(size - i - 1) Else pool(j) = pool(size - i - 1)
        Next i
    End If
    ArraySample = result
End Function

'=============================================================================
' LinSpace — 生成 n 个等间距点 (类似 numpy linspace)
'=============================================================================
Public Function LinSpace(ByVal start As Double, ByVal endVal As Double, ByVal n As Long) As Variant
    Dim result() As Double, i As Long
    If n < 2 Then LinSpace = Array(): Exit Function
    ReDim result(0 To n - 1)
    Dim stepVal As Double: stepVal = (endVal - start) / (n - 1)
    For i = 0 To n - 1: result(i) = start + i * stepVal: Next i
    LinSpace = result
End Function

'=============================================================================
' RangeFill — 生成等差数列 (输出 0-based)
'=============================================================================
Public Function RangeFill( _
    ByVal start As Double, _
    ByVal count As Long, _
    Optional ByVal stepSize As Double = 1#) As Variant

    Dim result() As Double
    Dim i As Long

    If count <= 0 Then
        RangeFill = Array()
        Exit Function
    End If

    ReDim result(0 To count - 1)
    For i = 0 To count - 1
        result(i) = start + i * stepSize
    Next i
    RangeFill = result
End Function

'=============================================================================
' ArrayChunk — 将一维数组按块分组（返回二维 Variant 数组，每行一个块）
'
' 输出: (0 To chunkCount-1, 0 To size-1)，不足的用 Empty 填充
' 注意: VBA 不支持 0 元素维度的数组，因此空输入返回一维 Array()
'       而非二维数组。调用方应先用 ArrayDims 或 UBound 检查结果。
'=============================================================================
Public Function ArrayChunk( _
    ByRef arr As Variant, _
    ByVal size As Long) As Variant

    Dim result() As Variant
    Dim lb As Long, ub As Long, n As Long
    Dim chunkCount As Long
    Dim i As Long, chunk As Long, col As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayChunk = data
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayChunk", "需要一维数组。"

    If size < 1 Then
        Err.Raise ERR_INVALID_INPUT, "ArrayChunk", "chunkSize 必须为正数。"
        Exit Function
    End If

    lb = LBound(data): ub = UBound(data)
    n = ub - lb + 1
    If n = 0 Then
        ' 空数组 — 按 SKILL.md 约定返回空 Variant 数组
        ArrayChunk = Array()
        Exit Function
    End If
    chunkCount = (n + size - 1) \ size

    ReDim result(0 To chunkCount - 1, 0 To size - 1)
    For i = lb To ub
        chunk = (i - lb) \ size
        col = (i - lb) Mod size
        result(chunk, col) = data(i)
    Next i
    ArrayChunk = result
End Function

'=============================================================================

'=============================================================================
' ArrayMin / ArrayMax / ArraySum — 数值聚合
'
' ArrayMin 和 ArrayMax 返回 Variant:
'   - 找到数值时返回 Double
'   - 无数值元素时返回 Empty
' ArraySum 返回 Variant（一维数组返回数值总和，二维数组返回 CVErr；无效输入返回 Empty）
'=============================================================================
Public Function ArrayMin(ByRef arr As Variant) As Variant
    Dim data As Variant
    Dim lb As Long, ub As Long, i As Long
    Dim result As Double, currVal As Double
    Dim found As Boolean

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayMin = Empty
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayMin", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    For i = lb To ub
        If Not (VarType(data(i)) = vbBoolean) And IsNumeric(data(i)) And Not IsEmpty(data(i)) And Not IsNull(data(i)) Then
            currVal = CDbl(data(i))
            If Not found Then
                result = currVal
                found = True
            ElseIf currVal < result Then
                result = currVal
            End If
        End If
    Next i
    If found Then ArrayMin = result Else ArrayMin = Empty
End Function

Public Function ArrayMax(ByRef arr As Variant) As Variant
    Dim data As Variant
    Dim lb As Long, ub As Long, i As Long
    Dim result As Double, currVal As Double
    Dim found As Boolean

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayMax = Empty
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayMax", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    For i = lb To ub
        If Not (VarType(data(i)) = vbBoolean) And IsNumeric(data(i)) And Not IsEmpty(data(i)) And Not IsNull(data(i)) Then
            currVal = CDbl(data(i))
            If Not found Then
                result = currVal
                found = True
            ElseIf currVal > result Then
                result = currVal
            End If
        End If
    Next i
    If found Then ArrayMax = result Else ArrayMax = Empty
End Function

Public Function ArraySum(ByRef arr As Variant) As Variant
    Dim data As Variant
    Dim lb As Long, ub As Long, i As Long
    Dim sum As Double, c As Double, y As Double, t As Double

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArraySum = 0#
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArraySum", "需要一维数组。"

    sum = 0#: c = 0#
    lb = LBound(data): ub = UBound(data)
    For i = lb To ub
        If Not (VarType(data(i)) = vbBoolean) And IsNumeric(data(i)) And Not IsEmpty(data(i)) And Not IsNull(data(i)) Then
            y = CDbl(data(i)) - c
            t = sum + y
            c = (t - sum) - y
            sum = t
        End If
    Next i
    ArraySum = sum
End Function

'=============================================================================
' ArrayToString — 数组转字符串表示
'=============================================================================
Public Function ArrayToString( _
    ByRef arr As Variant, _
    Optional ByVal delimiter As String = ", ") As Variant

    Dim lb As Long, ub As Long, i As Long
    Dim parts() As String
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayToString = SafeStr(data)
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayToString", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    If ub < lb Then
        ArrayToString = ""
        Exit Function
    End If
    ReDim parts(lb To ub)
    For i = lb To ub: parts(i) = SafeStr(data(i)): Next i
    ArrayToString = Join(parts, delimiter)
End Function

'=============================================================================
' ArrayReverse — 反转一维数组 (输出 0-based)
'=============================================================================
Public Function ArrayReverse(ByRef arr As Variant) As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim lb As Long, ub As Long, n As Long
    Dim i As Long

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayReverse = data
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayReverse", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    If ub < lb Then
        ArrayReverse = Array()
        Exit Function
    End If
    n = ub - lb
    ReDim result(0 To n)
    For i = lb To ub
        result(n - (i - lb)) = data(i)
    Next i
    ArrayReverse = result
End Function

'=============================================================================

'=============================================================================

'=============================================================================
' ArrayGetRow — 提取二维数组的指定行（输出 0-based）
'
' row: 0-based 行索引（即第 0 行为第一行）
' 越界时引发错误
'=============================================================================
Public Function ArrayGetRow(ByRef arr As Variant, ByVal row As Long) As Variant
    Dim result() As Variant
    Dim nCols As Long, j As Long
    Dim r0 As Long, c0 As Long, nRows As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        Err.Raise ERR_OUT_OF_BOUNDS, "ArrayGetRow", "rowIndex 越界。"
        Exit Function
    End If

    If ArrayDims(data) <> 2 Then
        Err.Raise ERR_OUT_OF_BOUNDS, "ArrayGetRow", "rowIndex 越界。"
        Exit Function
    End If

    r0 = LBound(data, 1): c0 = LBound(data, 2)
    nRows = UBound(data, 1) - r0 + 1
    nCols = UBound(data, 2) - c0 + 1

    If row < 0 Or row >= nRows Then
        Err.Raise ERR_OUT_OF_BOUNDS, "ArrayGetRow", "rowIndex 越界。"
        Exit Function
    End If

    ReDim result(0 To nCols - 1)
    For j = 1 To nCols
        result(j - 1) = data(r0 + row, c0 + j - 1)
    Next j
    ArrayGetRow = result
End Function

'=============================================================================
' ArrayGetCol — 提取二维数组的指定列（输出 0-based）
'
' col: 0-based 列索引（即第 0 列为第一列）
' 越界时引发错误
'=============================================================================
Public Function ArrayGetCol(ByRef arr As Variant, ByVal col As Long) As Variant
    Dim result() As Variant
    Dim nRows As Long, nCols As Long, i As Long
    Dim r0 As Long, c0 As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        Err.Raise ERR_OUT_OF_BOUNDS, "ArrayGetCol", "colIndex 越界。"
        Exit Function
    End If

    If ArrayDims(data) <> 2 Then
        Err.Raise ERR_OUT_OF_BOUNDS, "ArrayGetCol", "colIndex 越界。"
        Exit Function
    End If

    r0 = LBound(data, 1): c0 = LBound(data, 2)
    nRows = UBound(data, 1) - r0 + 1
    nCols = UBound(data, 2) - c0 + 1

    If col < 0 Or col >= nCols Then
        Err.Raise ERR_OUT_OF_BOUNDS, "ArrayGetCol", "colIndex 越界。"
        Exit Function
    End If

    ReDim result(0 To nRows - 1)
    For i = 1 To nRows
        result(i - 1) = data(r0 + i - 1, c0 + col)
    Next i
    ArrayGetCol = result
End Function

'=============================================================================
' ArrayTranspose2D — 二维数组行列转置 (输出 1-based, 遵循 VBA Range 惯例)
'=============================================================================
Public Function ArrayTranspose2D(ByRef arr As Variant) As Variant
    Dim result() As Variant
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim r0 As Long, c0 As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayTranspose2D = data
        Exit Function
    End If

    If IsArray1D(data) Then
        ArrayTranspose2D = ArrayTranspose1D(data, True)
        Exit Function
    End If

    If ArrayDims(data) <> 2 Then
        Err.Raise ERR_2D_REQUIRED, "ArrayTranspose2D", "需要二维数组。"
        Exit Function
    End If

    r0 = LBound(data, 1): c0 = LBound(data, 2)
    nRows = UBound(data, 1) - r0 + 1
    nCols = UBound(data, 2) - c0 + 1

    ReDim result(1 To nCols, 1 To nRows)
    For i = 1 To nRows
        For j = 1 To nCols
            result(j, i) = data(r0 + i - 1, c0 + j - 1)
        Next j
    Next i
    ArrayTranspose2D = result
End Function

'=============================================================================
' ArrayEqual — 元素逐一比较
'=============================================================================
Public Function ArrayEqual( _
    ByRef arr1 As Variant, _
    ByRef arr2 As Variant, _
    Optional ByVal caseSensitive As Boolean = False) As Variant

    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    Dim i As Long
    Dim compare As VbCompareMethod
    Dim data1 As Variant, data2 As Variant

    data1 = NormalizeToArray(arr1)
    data2 = NormalizeToArray(arr2)
    If Not IsArray(data1) Or Not IsArray(data2) Then
        ArrayEqual = FilterEquals(data1, data2)
        Exit Function
    End If

    If Not CheckIs1D(data1) Then Err.Raise ERR_1D_REQUIRED, "ArrayEqual", "需要一维数组。"
    If Not CheckIs1D(data2) Then Err.Raise ERR_1D_REQUIRED, "ArrayEqual", "需要一维数组。"

    lb1 = LBound(data1): ub1 = UBound(data1)
    lb2 = LBound(data2): ub2 = UBound(data2)

    If ub1 - lb1 <> ub2 - lb2 Then
        ArrayEqual = False
        Exit Function
    End If

    If caseSensitive Then compare = vbBinaryCompare Else compare = vbTextCompare

    For i = 0 To ub1 - lb1
        If IsNull(data1(lb1 + i)) And IsNull(data2(lb2 + i)) Then
            ' 相等 — 继续
        ElseIf IsNull(data1(lb1 + i)) Or IsNull(data2(lb2 + i)) Then
            ArrayEqual = False
            Exit Function
        ElseIf IsError(data1(lb1 + i)) And IsError(data2(lb2 + i)) Then
            ' 相等 — 继续
        ElseIf IsError(data1(lb1 + i)) Or IsError(data2(lb2 + i)) Then
            ArrayEqual = False
            Exit Function
        ElseIf VarType(data1(lb1 + i)) <> vbBoolean And VarType(data2(lb2 + i)) <> vbBoolean And _
                IsNumeric(data1(lb1 + i)) And IsNumeric(data2(lb2 + i)) Then
            If Abs(CDbl(data1(lb1 + i)) - CDbl(data2(lb2 + i))) >= 1E-12 Then
                ArrayEqual = False
                Exit Function
            End If
        Else
            If StrComp(SafeStr(data1(lb1 + i)), SafeStr(data2(lb2 + i)), compare) <> 0 Then
                ArrayEqual = False
                Exit Function
            End If
        End If
    Next i
    ArrayEqual = True
End Function

'=============================================================================
' ArrayProduct — 数值元素连乘积
'
' 当且仅当所有元素被排除时返回 0（与 Excel PRODUCT 工作表函数行为一致；
' 数学上空积应为 1，但本库以 Excel 兼容性优先）。零值参与乘法时 prod=0。
'=============================================================================
Public Function ArrayProduct(ByRef arr As Variant) As Variant
    Dim lb As Long, ub As Long, i As Long
    Dim prod As Double
    Dim found As Boolean
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ' 标量：数值返回自身，非数值返回 0（与 Excel PRODUCT 一致）
        If VarType(data) <> vbBoolean And IsNumeric(data) Then ArrayProduct = CDbl(data) Else ArrayProduct = 0#
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayProduct", "需要一维数组。"

    prod = 1#
    lb = LBound(data): ub = UBound(data)
    For i = lb To ub
        If Not (VarType(data(i)) = vbBoolean) And IsNumeric(data(i)) And Not IsEmpty(data(i)) And Not IsNull(data(i)) Then
            prod = prod * CDbl(data(i))
            found = True
        End If
    Next i
    ' 无有效数值时返回 0（与 Excel PRODUCT 一致，而非数学空积 1）
    If found Then ArrayProduct = prod Else ArrayProduct = 0#
End Function

'=============================================================================
' CumSum — 累积和（输出 0-based）
'
' 返回与输入等长的一维数组，第 i 个元素 = arr(0..i) 之和
'=============================================================================
Public Function CumSum(ByRef arr As Variant) As Variant
    Dim result() As Double
    Dim lb As Long, ub As Long, n As Long
    Dim i As Long, idx As Long, total As Double
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        CumSum = data
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "CumSum", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    If ub < lb Then
        CumSum = Array()
        Exit Function
    End If
    n = ub - lb
    ReDim result(0 To n)
    total = 0#
    idx = 0
    For i = lb To ub
        If Not (VarType(data(i)) = vbBoolean) And IsNumeric(data(i)) And Not IsEmpty(data(i)) And Not IsNull(data(i)) Then total = total + CDbl(data(i))
        result(idx) = total
        idx = idx + 1
    Next i
    CumSum = result
End Function

'=============================================================================
' ArgSort — 返回排序后的索引（输出 0-based）
'
' 返回: 一维 Long 数组，其中 result[0] 是最小值的索引
' 用于：获取排名、按某列对其他列排序等
'=============================================================================
Public Function ArgSort( _
    ByRef arr As Variant, _
    Optional ByVal ascending As Boolean = True) As Variant

    Dim result() As Long
    Dim lb As Long, ub As Long, n As Long
    Dim i As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArgSort = data
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArgSort", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    n = ub - lb + 1
    If n = 0 Then
        ArgSort = Array()
        Exit Function
    End If

    ReDim result(0 To n - 1)
    For i = 0 To n - 1
        result(i) = lb + i
    Next i

    AO.SortIndices data, result, ascending
    ArgSort = result
End Function

'=============================================================================
' ArrayAny — 是否存在任意元素满足条件
'=============================================================================
Public Function ArrayAny( _
    ByRef arr As Variant, _
    ByVal matchValue As Variant, _
    Optional ByVal operator As String = "=") As Variant

    Dim lb As Long, ub As Long, i As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayAny = FilterPasses(data, matchValue, operator)
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayAny", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    For i = lb To ub
        If FilterPasses(data(i), matchValue, operator) Then
            ArrayAny = True
            Exit Function
        End If
    Next i
    ArrayAny = False
End Function

'=============================================================================
' ArrayAll — 是否所有元素都满足条件
'=============================================================================
Public Function ArrayAll( _
    ByRef arr As Variant, _
    ByVal matchValue As Variant, _
    Optional ByVal operator As String = "=") As Variant

    Dim lb As Long, ub As Long, i As Long
    Dim data As Variant

    data = NormalizeToArray(arr)
    If Not IsArray(data) Then
        ArrayAll = FilterPasses(data, matchValue, operator)
        Exit Function
    End If

    If Not CheckIs1D(data) Then Err.Raise ERR_1D_REQUIRED, "ArrayAll", "需要一维数组。"

    lb = LBound(data): ub = UBound(data)
    ' 空数组：vacuously true（所有零个元素都满足条件）
    If ub < lb Then
        ArrayAll = True
        Exit Function
    End If
    For i = lb To ub
        If Not FilterPasses(data(i), matchValue, operator) Then
            ArrayAll = False
            Exit Function
        End If
    Next i
    ArrayAll = True
End Function

'=============================================================================
' 工作表函数 (UDF_ARR_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_ARR_UNIQUE(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_UNIQUE = ArrayUnique(arr): Exit Function
EH: UDF_ARR_UNIQUE = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_SORT(ByVal arr As Variant, Optional ByVal ascending As Variant = True) As Variant
    On Error GoTo EH: UDF_ARR_SORT = ArraySort(arr, ascending): Exit Function
EH: UDF_ARR_SORT = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_FILTER(ByVal arr As Variant, ByVal matchValue As Variant, Optional ByVal operator As Variant = "=") As Variant
    On Error GoTo EH: UDF_ARR_FILTER = ArrayFilterByValue(arr, matchValue, operator): Exit Function
EH: UDF_ARR_FILTER = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_CONCAT(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    On Error GoTo EH: UDF_ARR_CONCAT = ArrayConcat(arr1, arr2): Exit Function
EH: UDF_ARR_CONCAT = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_SLICE(ByVal arr As Variant, Optional ByVal start As Variant = 0, Optional ByVal cnt As Variant = -1) As Variant
    On Error GoTo EH: UDF_ARR_SLICE = ArraySlice(arr, start, cnt): Exit Function
EH: UDF_ARR_SLICE = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_FLATTEN(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_FLATTEN = ArrayFlatten(arr): Exit Function
EH: UDF_ARR_FLATTEN = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_TRANSPOSE(ByVal arr As Variant, Optional ByVal asColumn As Variant = True) As Variant
    On Error GoTo EH: UDF_ARR_TRANSPOSE = ArrayTranspose1D(arr, asColumn): Exit Function
EH: UDF_ARR_TRANSPOSE = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_FIND(ByVal arr As Variant, ByVal value As Variant, Optional ByVal caseSensitive As Variant = False) As Variant
    On Error GoTo EH: UDF_ARR_FIND = ArrayFind(arr, value, caseSensitive): Exit Function
EH: UDF_ARR_FIND = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_CONTAINS(ByVal arr As Variant, ByVal value As Variant, Optional ByVal caseSensitive As Variant = False) As Variant
    On Error GoTo EH: UDF_ARR_CONTAINS = ArrayContains(arr, value, caseSensitive): Exit Function
EH: UDF_ARR_CONTAINS = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_LOOKUP(ByVal lookupArray As Variant, ByVal lookupValue As Variant, _
    ByVal lookupCol As Variant, Optional ByVal returnCols As Variant, _
    Optional ByVal matchType As Variant = 0) As Variant
    On Error GoTo EH: UDF_ARR_LOOKUP = ArrayLookup(lookupArray, lookupValue, lookupCol, returnCols, matchType): Exit Function
EH: UDF_ARR_LOOKUP = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_LINSPACE(ByVal start As Variant, ByVal endVal As Variant, ByVal n As Variant) As Variant
    On Error GoTo EH: UDF_ARR_LINSPACE = LinSpace(start, endVal, n): Exit Function
EH: UDF_ARR_LINSPACE = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_SHUFFLE(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_SHUFFLE = ArrayShuffle(arr): Exit Function
EH: UDF_ARR_SHUFFLE = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_RANGEFILL(ByVal start As Variant, ByVal count As Variant, Optional ByVal stepSize As Variant = 1#) As Variant
    On Error GoTo EH: UDF_ARR_RANGEFILL = RangeFill(start, count, stepSize): Exit Function
EH: UDF_ARR_RANGEFILL = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_CHUNK(ByVal arr As Variant, ByVal size As Variant) As Variant
    On Error GoTo EH: UDF_ARR_CHUNK = ArrayChunk(arr, size): Exit Function
EH: UDF_ARR_CHUNK = CVErr(xlErrValue)
End Function


Public Function UDF_ARR_MIN(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_MIN = ArrayMin(arr): Exit Function
EH: UDF_ARR_MIN = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_MAX(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_MAX = ArrayMax(arr): Exit Function
EH: UDF_ARR_MAX = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_SUM(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_SUM = ArraySum(arr): Exit Function
EH: UDF_ARR_SUM = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_TOSTRING(ByVal arr As Variant, Optional ByVal delimiter As Variant = ", ") As Variant
    On Error GoTo EH: UDF_ARR_TOSTRING = ArrayToString(arr, delimiter): Exit Function
EH: UDF_ARR_TOSTRING = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_REVERSE(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_REVERSE = ArrayReverse(arr): Exit Function
EH: UDF_ARR_REVERSE = CVErr(xlErrValue)
End Function



Public Function UDF_ARR_GETROW(ByVal arr As Variant, ByVal row As Variant) As Variant
    On Error GoTo EH: UDF_ARR_GETROW = ArrayGetRow(arr, row): Exit Function
EH: UDF_ARR_GETROW = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_GETCOL(ByVal arr As Variant, ByVal col As Variant) As Variant
    On Error GoTo EH: UDF_ARR_GETCOL = ArrayGetCol(arr, col): Exit Function
EH: UDF_ARR_GETCOL = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_TRANSPOSE2D(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_TRANSPOSE2D = ArrayTranspose2D(arr): Exit Function
EH: UDF_ARR_TRANSPOSE2D = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_EQUAL(ByVal arr1 As Variant, ByVal arr2 As Variant, Optional ByVal caseSensitive As Variant = False) As Variant
    On Error GoTo EH: UDF_ARR_EQUAL = ArrayEqual(arr1, arr2, caseSensitive): Exit Function
EH: UDF_ARR_EQUAL = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_PRODUCT(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_PRODUCT = ArrayProduct(arr): Exit Function
EH: UDF_ARR_PRODUCT = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_CUMSUM(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_ARR_CUMSUM = CumSum(arr): Exit Function
EH: UDF_ARR_CUMSUM = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_ARGSORT(ByVal arr As Variant, Optional ByVal ascending As Variant = True) As Variant
    On Error GoTo EH: UDF_ARR_ARGSORT = ArgSort(arr, ascending): Exit Function
EH: UDF_ARR_ARGSORT = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_ANY(ByVal arr As Variant, ByVal matchValue As Variant, Optional ByVal operator As Variant = "=") As Variant
    On Error GoTo EH: UDF_ARR_ANY = ArrayAny(arr, matchValue, operator): Exit Function
EH: UDF_ARR_ANY = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_ALL(ByVal arr As Variant, ByVal matchValue As Variant, Optional ByVal operator As Variant = "=") As Variant
    On Error GoTo EH: UDF_ARR_ALL = ArrayAll(arr, matchValue, operator): Exit Function
EH: UDF_ARR_ALL = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_COUNTIF(ByVal arr As Variant, ByVal matchValue As Variant, Optional ByVal operator As Variant = "=") As Variant
    On Error GoTo EH: UDF_ARR_COUNTIF = ArrayCountIf(arr, matchValue, operator): Exit Function
EH: UDF_ARR_COUNTIF = CVErr(xlErrValue)
End Function

Public Function UDF_ARR_SAMPLE(ByVal arr As Variant, ByVal n As Variant, Optional ByVal withReplacement As Variant = False) As Variant
    On Error GoTo EH: UDF_ARR_SAMPLE = ArraySample(arr, n, withReplacement): Exit Function
EH: UDF_ARR_SAMPLE = CVErr(xlErrValue)
End Function