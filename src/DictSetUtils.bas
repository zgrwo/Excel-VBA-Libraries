Option Explicit

'==============================================================================
' Module:       DictSetUtils
' Purpose:      Dictionary/Set operations: merge, intersect, difference, frequency
' Layer:        Data
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       36 functions/subs
'==============================================================================

Private DP As New DictProxy
Private VK As New VariantKit

'=====================================================================
' DictSetUtils.bas — 字典工具 & 集合运算
'
' 错误码:
'   vbObjectError + 1101 — 无效输入 / 外部组件不可用
'
' 字典操作:
'   DictMerge         — 合并两个字典
'   DictInvert        — 键值互换
'   DictKeys          — 键 → 数组
'   DictValues        — 值 → 数组
'   DictTo2DArray     — 字典 → 二维数组 (键值对)
'   ArrayToDict       — 数组 → 字典
'   CountFrequency    — 频数统计
'   GroupCount        — 频数统计直接输出 2D 数组
'   DictPick          — 按指定键列表提取子集
'   DictFrom2DArray   — 二维数组的某两列 → 字典
'   DictCount         — 字典条目数
'   DictIsEmpty       — 是否为空
'   DictClone         — 浅克隆
'   DictGetDefault    — 安全取值 (键不存在返回默认)
'   DictRenameKey     — 重命名键
'   DictRemoveKeys    — 批量删除键
'   DictMergeSum      — 合并字典，同键值求和
'   DictFilterByValue — 按值条件筛选子集
'   DictSortByKey     — 按键排序 → 2D 数组
'   DictSortByValue   — 按值排序 → 2D 数组
'   DictTopN          — 按值取前 N 条
'
' 工作表函数 (UDF_DICT_*): UNION, INTERSECT, DIFFERENCE, SYM_DIFF, CARTESIAN, ISSUBSET, ISEQUAL
'
'   SetUnion           — 并集
'   SetIntersect       — 交集
'   SetDifference      — 差集 (A - B)
'   SetSymDifference   — 对称差
'   SetIsSubset        — 子集判断
'   SetEqual           — 判断两集合相等
'   SetCartesianProduct — 笛卡尔积
'=====================================================================

'=============================================================================
' 字典操作
'=============================================================================

' -- 错误代码常量 -----------------------------------------------------
Private Const ERR_INVALID_INPUT As Long = vbObjectError + 1101


' -- DictKey — 生成安全的字典键，避免 Null/Error/Object/Empty/Array 碰撞到空字符串 --
' 数值通过 CDbl 归一化以确保 Integer/Long/Double/Single 同一值映射到相同键
' 注: 与 VariantKit.SafeKey 语义相近 (SafeKey→ArrayToKey 已有完整逐元素数组键消歧)，
'     但保留本模块键格式契约 (CountFrequency 等向用户暴露这些键)，
'     故未直接委托 (§17.1 已知技术债 — 委托会改变用户可见键格式；
'     长期方案: VK.SafeKey 格式可参数化后迁移)。
Private Function DictKey(ByVal v As Variant) As String
    If IsNull(v) Then
        DictKey = "##NULL##"
    ElseIf IsError(v) Then
        DictKey = "##ERROR##"
    ElseIf IsObject(v) Then
        DictKey = "##OBJECT:" & Hex$(ObjPtr(v)) & "##"
    ElseIf IsEmpty(v) Then
        DictKey = "##EMPTY##"
    ElseIf IsArray(v) Then
        DictKey = "##ARRAY##"  ' 数组无法 CStr — 独立哨兵键 (修复原 CStr(array) → Error 13)
    ElseIf IsNumeric(v) Then
        If VarType(v) = vbCurrency Then
            DictKey = CStr(v)  ' vbCurrency: exact fixed-point, avoid CDbl precision loss
        Else
            DictKey = CStr(CDbl(v))
        End If
    Else
        DictKey = CStr(v)
    End If
End Function

' -- InitDictPreserveMode — 创建新字典并继承源字典的 CompareMode --
Private Sub InitDictPreserveMode(ByRef result As Object, ByRef source As Object)
    Set result = DP.Create()
    If Not source Is Nothing Then
        ' 保留任意 CompareMode (非仅 vbTextCompare)
        result.CompareMode = source.CompareMode
    End If
End Sub

' DictMerge — 合并两个字典 (重复键: 默认保留 dict1 的值)
' 若 overwrite=True, 用 dict2 的值覆盖
Public Function DictMerge( _
    ByRef dict1 As Object, _
    ByRef dict2 As Object, _
    Optional ByVal overwrite As Boolean = False) As Object

    Dim result As Object
    Dim key As Variant

    InitDictPreserveMode result, dict1

    If Not dict1 Is Nothing Then
        For Each key In dict1.Keys
            result.Add key, dict1(key)
        Next key
    End If

    If Not dict2 Is Nothing Then
        For Each key In dict2.Keys
            If result.Exists(key) Then
                If overwrite Then
                    result.Remove key
                    result.Add key, dict2(key)
                End If
            Else
                result.Add key, dict2(key)
            End If
        Next key
    End If

    Set DictMerge = result
End Function

' DictMergeSum — 合并字典，同键数值求和
Public Function DictMergeSum( _
    ByRef dict1 As Object, _
    ByRef dict2 As Object) As Object

    Dim result As Object
    Dim key As Variant

    InitDictPreserveMode result, dict1

    If Not dict1 Is Nothing Then
        For Each key In dict1.Keys
            result.Add key, dict1(key)
        Next key
    End If

    If Not dict2 Is Nothing Then
        For Each key In dict2.Keys
            If result.Exists(key) Then
                ' 安全提取值 — IsObject 派发避免对象默认属性求值 (§3.2/§5.2.2)
                ' 先提取再决策: 消除重复的 Remove+Add 模式 + 减少字典查找次数
                Dim rVal As Variant, d2Val As Variant
                If IsObject(result(key)) Then Set rVal = result(key) Else rVal = result(key)
                If IsObject(dict2(key)) Then Set d2Val = dict2(key) Else d2Val = dict2(key)

                Dim canSum As Boolean
                canSum = False
                If Not IsObject(rVal) And Not IsObject(d2Val) Then
                    canSum = (VarType(rVal) <> vbBoolean And VarType(d2Val) <> vbBoolean And _
                              IsNumeric(rVal) And IsNumeric(d2Val))
                End If

                result.Remove key
                If canSum Then
                    result.Add key, CDbl(rVal) + CDbl(d2Val)
                Else
                    ' 对象值 / 非数值 / 布尔 → dict2 覆盖
                    result.Add key, d2Val
                End If
            Else
                result.Add key, dict2(key)
            End If
        Next key
    End If

    Set DictMergeSum = result
End Function

' DictInvert — 键值互换 (重复值后出现的键覆盖先出现的; 跳过对象/错误/Null 值, 它们无法作为字典键)
Public Function DictInvert(ByRef dict As Object) As Object
    Dim result As Object
    Dim key As Variant

    InitDictPreserveMode result, dict
    If dict Is Nothing Then
        Set DictInvert = result
        Exit Function
    End If

    Dim newKey As String
    For Each key In dict.Keys
        If IsObject(dict(key)) Then
            ' 跳过对象类型值，无法作为字典键
        ElseIf IsError(dict(key)) Then
            ' 跳过错误类型值，无法作为字典键
        ElseIf IsNull(dict(key)) Then
            ' 跳过 Null 值，无法作为字典键
        ElseIf IsArray(dict(key)) Then
            ' 跳过数组值，无法作为字典键
        Else
            ' .Add 模式 — 避免对含默认属性的对象触发隐式求值 (§5.2.2)
            newKey = CStr(dict(key))
            If result.Exists(newKey) Then result.Remove newKey
            result.Add newKey, key
        End If
    Next key
    Set DictInvert = result
End Function

' DictKeys — 返回键的 Variant 数组 (保持原始类型)
Public Function DictKeys(ByRef dict As Object) As Variant
    If dict Is Nothing Then
        DictKeys = Array()
        Exit Function
    End If
    DictKeys = dict.Keys
End Function

' DictValues — 返回值的 Variant 数组
Public Function DictValues(ByRef dict As Object) As Variant
    If dict Is Nothing Then
        DictValues = Array()
        Exit Function
    End If
    DictValues = dict.Items
End Function

' DictTo2DArray — 字典 → 二维 Variant 数组 (1 To n, 1 To 2)
Public Function DictTo2DArray(ByRef dict As Object) As Variant
    Dim result() As Variant
    Dim keys As Variant, items As Variant
    Dim i As Long, n As Long

    If dict Is Nothing Then DictTo2DArray = Array(): Exit Function
    If dict.Count = 0 Then DictTo2DArray = Array(): Exit Function

    keys = dict.Keys
    items = dict.Items
    n = dict.Count
    ReDim result(1 To n, 1 To 2)

    For i = 0 To n - 1
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = items(i)
    Next i
    DictTo2DArray = result
End Function

' ArrayToDict — 一维数组 → 字典
' 若 keyFn = "index" 则以索引为键
' 否则以元素值为键 (重复键保留第一个)
Public Function ArrayToDict( _
    ByRef arr As Variant, _
    Optional ByVal keyFn As String = "value") As Object

    Dim result As Object
    Dim i As Long, lb As Long, ub As Long
    Set result = DP.Create()

    If Not IsArray(arr) Then
        Set ArrayToDict = result
        Exit Function
    End If

    Dim k As String
    lb = LBound(arr): ub = UBound(arr)
    For i = lb To ub
        If LCase$(keyFn) = "index" Then
            result.Add CStr(i), arr(i)
        Else
            k = DictKey(arr(i))
            If Not result.Exists(k) Then
                result.Add k, arr(i)
            End If
        End If
    Next i
    Set ArrayToDict = result
End Function

' DictFrom2DArray — 二维数组的某两列 → 字典 (colKey → colValue)
Public Function DictFrom2DArray( _
    ByRef arr As Variant, _
    Optional ByVal colKey As Long = 1, _
    Optional ByVal colValue As Long = 2) As Object

    Dim result As Object
    Dim nRows As Long, i As Long
    Dim r0 As Long, c0 As Long
    Dim nCols As Long
    Dim key As String
    Dim cellVal As Variant

    Set result = DP.Create()

    If Not IsArray(arr) Then
        Set DictFrom2DArray = result
        Exit Function
    End If

    ' 检查是否为 2D 数组
    Err.Clear
    On Error Resume Next
    Dim dummy As Long
    dummy = UBound(arr, 2)
    If Err.Number <> 0 Then
        ' 1D 数组 → 索引为键
        Err.Clear: On Error GoTo 0
        Set DictFrom2DArray = ArrayToDict(arr, "index")
        Exit Function
    End If
    Err.Clear: On Error GoTo 0

    ' 拒绝 3D+ 数组
    On Error Resume Next
    dummy = UBound(arr, 3)
    If Err.Number = 0 Then
        Err.Clear: On Error GoTo 0
        Set DictFrom2DArray = result
        Exit Function
    End If
    Err.Clear: On Error GoTo 0

    r0 = LBound(arr, 1): c0 = LBound(arr, 2)
    nRows = UBound(arr, 1) - r0 + 1
    nCols = UBound(arr, 2) - c0 + 1

    If colKey < 1 Or colKey > nCols Or colValue < 1 Or colValue > nCols Then
        Set DictFrom2DArray = result
        Exit Function
    End If

    For i = 1 To nRows
        cellVal = arr(r0 + i - 1, c0 + colKey - 1)
        key = DictKey(cellVal)
        If Not result.Exists(key) Then
            result.Add key, arr(r0 + i - 1, c0 + colValue - 1)
        End If
    Next i
    Set DictFrom2DArray = result
End Function

' CountFrequency — 频数统计 (返回: key → count 字典)
Public Function CountFrequency(ByVal arr As Variant) As Object
    Dim result As Object
    Dim i As Long, lb As Long, ub As Long
    Dim key As String

    VK.NormalizeInput arr, True  ' Range→Array + 2D→1D

    ' 先初始化 — 确保非数组早退路径返回空字典而非 Nothing (修复 H3)。
    ' 在 IsArray 之前分配是故意的: 标量输入 (如单单元格 Range→标量) 需返回空字典，
    ' 而非 Nothing (调用方 .Exists() 会 Error 91)。COM 分配开销在冷路径上可接受。
    Set result = DP.Create()
    If Not IsArray(arr) Then
        Set CountFrequency = result
        Exit Function
    End If

    lb = LBound(arr): ub = UBound(arr)
    For i = lb To ub
        key = DictKey(arr(i))
        If result.Exists(key) Then
            result(key) = CLng(result(key)) + 1
        Else
            result.Add key, 1
        End If
    Next i
    Set CountFrequency = result
End Function

'=============================================================================
' GroupCount — 频数统计直接输出 2D 数组 (返回: (key, count) 对)
' 比 CountFrequency 更轻量，适合工作表直接使用
'=============================================================================
Public Function GroupCount(ByVal arr As Variant) As Variant
    Dim dict As Object, i As Long, lb As Long, ub As Long, ki As Long
    Dim data As Variant, result() As Variant, key As Variant

    VK.NormalizeInput arr, True  ' Range→Array + 2D→1D

    If IsArray(arr) Then data = arr Else Err.Raise ERR_INVALID_INPUT, "GroupCount", "需要数组输入。"
    lb = LBound(data): ub = UBound(data)
    Set dict = DP.Create()
    For i = lb To ub
        If Not IsError(data(i)) And Not IsNull(data(i)) Then
            key = DictKey(data(i))
            If dict.Exists(key) Then dict(key) = CLng(dict(key)) + 1 Else dict.Add key, 1
        End If
    Next i
    If dict.Count = 0 Then GroupCount = Array(): Exit Function
    ReDim result(1 To dict.Count, 1 To 2): ki = 1
    For Each key In dict.Keys: result(ki, 1) = key: result(ki, 2) = dict(key): ki = ki + 1: Next
    GroupCount = result
End Function

' DictPick — 按指定键列表提取子集
Public Function DictPick(ByRef dict As Object, ByVal keys As Variant) As Object
    Dim result As Object
    Dim i As Long, lb As Long, ub As Long
    Dim key As Variant

    InitDictPreserveMode result, dict

    If dict Is Nothing Then
        Set DictPick = result
        Exit Function
    End If
    If Not IsArray(keys) Then
        Set DictPick = result
        Exit Function
    End If

    lb = LBound(keys): ub = UBound(keys)
    For i = lb To ub
        key = keys(i)
        If dict.Exists(key) Then
            If Not result.Exists(key) Then result.Add key, dict(key)
        End If
    Next i
    Set DictPick = result
End Function

' DictCount — 字典条目数
Public Function DictCount(ByRef dict As Object) As Long
    If dict Is Nothing Then
        DictCount = 0
    Else
        DictCount = dict.Count
    End If
End Function

'=============================================================================
' 集合运算 (基于一维数组)
'=============================================================================

' 将数组转为字符串键集合 (内部使用)
Private Function ArrayToSet(ByRef arr As Variant) As Object
    Dim result As Object
    Dim i As Long, lb As Long, ub As Long
    Dim k As String

    Dim localArr As Variant: localArr = arr
    VK.NormalizeInput localArr, True  ' Range→Array + 2D→1D (local copy — §1.1)

    ' 先初始化 — 确保非数组早退路径返回空字典而非 Nothing (修复 H4)。
    ' 在 IsArray 之前分配是故意的: 标量输入需返回空字典而非 Nothing。
    Set result = DP.Create()
    If Not IsArray(localArr) Then
        Set ArrayToSet = result
        Exit Function
    End If

    lb = LBound(localArr): ub = UBound(localArr)
    For i = lb To ub
        k = DictKey(localArr(i))
        If Not result.Exists(k) Then
            result.Add k, localArr(i)
        End If
    Next i
    Set ArrayToSet = result
End Function

' SetUnion — 并集
Public Function SetUnion(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant
    Dim dict As Object
    Dim result() As Variant
    Dim i As Long, v As Variant
    Dim k As String, lb2 As Long, ub2 As Long

    Set dict = ArrayToSet(arr1)
    ' 将 arr2 归一化后合并到 dict（与 ArrayToSet 一致的归一化路径）
    Dim localArr2 As Variant: localArr2 = arr2
    VK.NormalizeInput localArr2, True
    If IsArray(localArr2) Then
        lb2 = LBound(localArr2): ub2 = UBound(localArr2)
        For i = lb2 To ub2
            k = DictKey(localArr2(i))
            If Not dict.Exists(k) Then
                dict.Add k, localArr2(i)
            End If
        Next i
    End If

    If dict.Count = 0 Then
        result = Array()
        SetUnion = result
        Exit Function
    End If

    ReDim result(0 To dict.Count - 1)
    i = 0
    For Each v In dict.Items
        result(i) = v
        i = i + 1
    Next v
    SetUnion = result
End Function

' SetIntersect — 交集
Public Function SetIntersect(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant
    Dim set1 As Object, set2 As Object
    Dim result() As Variant
    Dim key As Variant
    Dim i As Long

    Set set1 = ArrayToSet(arr1)
    Set set2 = ArrayToSet(arr2)

    If set1.Count = 0 Or set2.Count = 0 Then
        result = Array()
        SetIntersect = result
        Exit Function
    End If

    ReDim result(0 To set1.Count - 1)
    i = 0
    For Each key In set1.Keys
        If set2.Exists(key) Then
            result(i) = set1(key)
            i = i + 1
        End If
    Next key
    If i = 0 Then
        result = Array()
    Else
        ReDim Preserve result(0 To i - 1)
    End If
    SetIntersect = result
End Function

' SetDifference — 差集 (arr1 - arr2)
Public Function SetDifference(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant
    Dim set1 As Object, set2 As Object
    Dim result() As Variant
    Dim key As Variant
    Dim i As Long

    Set set1 = ArrayToSet(arr1)
    Set set2 = ArrayToSet(arr2)

    If set1.Count = 0 Then
        result = Array()
        SetDifference = result
        Exit Function
    End If

    ReDim result(0 To set1.Count - 1)
    i = 0
    For Each key In set1.Keys
        If Not set2.Exists(key) Then
            result(i) = set1(key)
            i = i + 1
        End If
    Next key
    If i = 0 Then
        result = Array()
    Else
        ReDim Preserve result(0 To i - 1)
    End If
    SetDifference = result
End Function

' SetSymDifference — 对称差 (A Δ B)
Public Function SetSymDifference(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant
    Dim diff1 As Variant, diff2 As Variant
    diff1 = SetDifference(arr1, arr2)
    diff2 = SetDifference(arr2, arr1)
    SetSymDifference = SetUnion(diff1, diff2)
End Function

' SetIsSubset — 判断 arr1 是否为 arr2 的子集
Public Function SetIsSubset(ByRef arr1 As Variant, ByRef arr2 As Variant) As Boolean
    Dim set1 As Object, set2 As Object
    Dim key As Variant

    Set set1 = ArrayToSet(arr1)
    Set set2 = ArrayToSet(arr2)

    If set1.Count = 0 Then
        SetIsSubset = True
        Exit Function
    End If
    If set2.Count = 0 Then
        SetIsSubset = False
        Exit Function
    End If

    For Each key In set1.Keys
        If Not set2.Exists(key) Then
            SetIsSubset = False
            Exit Function
        End If
    Next key
    SetIsSubset = True
End Function

' SetEqual — 判断两集合是否相等
Public Function SetEqual(ByRef arr1 As Variant, ByRef arr2 As Variant) As Boolean
    Dim set1 As Object, set2 As Object
    Dim key As Variant

    Set set1 = ArrayToSet(arr1)
    Set set2 = ArrayToSet(arr2)

    If set1.Count <> set2.Count Then
        SetEqual = False
        Exit Function
    End If

    For Each key In set1.Keys
        If Not set2.Exists(key) Then
            SetEqual = False
            Exit Function
        End If
    Next key
    SetEqual = True
End Function

'=============================================================================
' DictFilterByValue — 按值条件筛选子字典
'=============================================================================
Public Function DictFilterByValue( _
    ByRef dict As Object, _
    ByVal matchValue As Variant, _
    Optional ByVal operator As String = "=") As Object

    Dim result As Object
    Dim key As Variant
    InitDictPreserveMode result, dict

    If dict Is Nothing Then
        Set DictFilterByValue = result
        Exit Function
    End If

    For Each key In dict.Keys
        If FilterPasses(dict(key), matchValue, operator) Then
            result.Add key, dict(key)
        End If
    Next key
    Set DictFilterByValue = result
End Function


' FilterPasses — delegates to VariantKit.FilterPasses (single source of truth)
Private Function FilterPasses(ByVal element As Variant, ByVal matchValue As Variant, ByVal operator As String) As Boolean
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    FilterPasses = vk.FilterPasses(element, matchValue, operator)
End Function


'=============================================================================
' DictSortByKey — 按键排序，返回 2D 数组
'=============================================================================
Public Function DictSortByKey( _
    ByRef dict As Object, _
    Optional ByVal ascending As Boolean = True) As Variant

    Dim keys As Variant
    Dim n As Long, i As Long
    Dim result() As Variant
    Dim idxArr() As Long

    If dict Is Nothing Then
        DictSortByKey = Array()
        Exit Function
    End If
    If dict.Count = 0 Then
        DictSortByKey = Array()
        Exit Function
    End If

    keys = dict.Keys
    n = UBound(keys) - LBound(keys) + 1

    ' Build index array, sort by key value (preserves original key types)
    ReDim idxArr(0 To n - 1)
    For i = 0 To n - 1: idxArr(i) = i: Next i
    Dim sortAO As ArrayOps: Set sortAO = New ArrayOps: sortAO.SortIndices keys, idxArr, ascending

    ReDim result(1 To n, 1 To 2)
    For i = 0 To n - 1
        result(i + 1, 1) = keys(idxArr(i))
        result(i + 1, 2) = dict(keys(idxArr(i)))
    Next i
    DictSortByKey = result
End Function

'=============================================================================
' DictSortByValue — 按值排序，返回 2D 数组
'=============================================================================
Public Function DictSortByValue( _
    ByRef dict As Object, _
    Optional ByVal ascending As Boolean = True) As Variant

    Dim result() As Variant
    Dim keys As Variant, items As Variant
    Dim n As Long, i As Long
    Dim idxArr() As Long

    If dict Is Nothing Then
        DictSortByValue = Array()
        Exit Function
    End If
    If dict.Count = 0 Then
        DictSortByValue = Array()
        Exit Function
    End If

    keys = dict.Keys
    items = dict.Items
    n = UBound(keys) - LBound(keys) + 1

    ' 构建索引数组，按值排序
    ReDim idxArr(0 To n - 1)
    For i = 0 To n - 1: idxArr(i) = i: Next i
    Dim sortAO2 As ArrayOps: Set sortAO2 = New ArrayOps: sortAO2.SortIndices items, idxArr, ascending

    ReDim result(1 To n, 1 To 2)
    For i = 0 To n - 1
        result(i + 1, 1) = keys(idxArr(i))
        result(i + 1, 2) = items(idxArr(i))
    Next i
    DictSortByValue = result
End Function

'=============================================================================
' DictGetDefault — 安全取值 (键不存在时返回默认值)
'=============================================================================
Public Function DictGetDefault( _
    ByRef dict As Object, _
    ByVal key As Variant, _
    Optional ByVal defaultValue As Variant) As Variant

    If dict Is Nothing Then
        If Not IsMissing(defaultValue) Then
            DictGetDefault = defaultValue
        Else
            DictGetDefault = Empty
        End If
        Exit Function
    End If

    If dict.Exists(key) Then
        If IsObject(dict(key)) Then
            Set DictGetDefault = dict(key)
        Else
            DictGetDefault = dict(key)
        End If
    Else
        If Not IsMissing(defaultValue) Then
            DictGetDefault = defaultValue
        Else
            DictGetDefault = Empty
        End If
    End If
End Function

'=============================================================================
' DictRenameKey — 重命名键 (返回新字典, 保留原字典不变)
'=============================================================================
Public Function DictRenameKey(ByRef dict As Object, ByVal oldKey As Variant, ByVal newKey As Variant) As Object
    Dim result As Object

    If dict Is Nothing Then
        Set result = DP.Create()
        Set DictRenameKey = result
        Exit Function
    End If
    Set result = DictClone(dict)

    If Not result.Exists(oldKey) Then
        Set DictRenameKey = result
        Exit Function
    End If
    If result.Exists(newKey) Then
        Set DictRenameKey = result
        Exit Function
    End If

    Dim existingVal As Variant
    If IsObject(result(oldKey)) Then
        Set existingVal = result(oldKey)
    Else
        existingVal = result(oldKey)
    End If
    result.Remove oldKey
    result.Add newKey, existingVal
    Set DictRenameKey = result
End Function

'=============================================================================
' DictRemoveKeys — 批量删除键 (返回新字典, 保留原字典不变)
'=============================================================================
Public Function DictRemoveKeys(ByRef dict As Object, ByRef keys As Variant) As Object
    Dim result As Object

    If dict Is Nothing Then
        Set result = DP.Create()
        Set DictRemoveKeys = result
        Exit Function
    End If
    Set result = DictClone(dict)
    If Not IsArray(keys) Then
        Set DictRemoveKeys = result
        Exit Function
    End If

    Dim i As Long, lb As Long, ub As Long
    lb = LBound(keys): ub = UBound(keys)
    For i = lb To ub
        If result.Exists(keys(i)) Then result.Remove keys(i)
    Next i
    Set DictRemoveKeys = result
End Function

'=============================================================================
' DictClone — 浅克隆 (值如果是对象引用，克隆和原字典指向同一对象)
'=============================================================================
Public Function DictClone(ByRef dict As Object) As Object
    Dim result As Object
    Dim key As Variant
    InitDictPreserveMode result, dict

    If dict Is Nothing Then
        Set DictClone = result
        Exit Function
    End If

    For Each key In dict.Keys
        result.Add key, dict(key)
    Next key
    Set DictClone = result
End Function

'=============================================================================
' DictIsEmpty — 是否为空
'=============================================================================
Public Function DictIsEmpty(ByRef dict As Object) As Boolean
    If dict Is Nothing Then
        DictIsEmpty = True
        Exit Function
    End If
    DictIsEmpty = (dict.Count = 0)
End Function

'=============================================================================
' DictTopN — 按值取前 N 条
'=============================================================================
Public Function DictTopN( _
    ByRef dict As Object, _
    ByVal n As Long, _
    Optional ByVal ascending As Boolean = False) As Variant

    Dim result() As Variant
    Dim allSorted As Variant
    Dim i As Long, cnt As Long

    allSorted = DictSortByValue(dict, ascending)
    ' 处理空字典 (DictSortByValue 返回 Array() — 空 Variant 数组)
    On Error Resume Next
    cnt = UBound(allSorted, 1)
    If Err.Number <> 0 Then  ' 非数组 (错误值) 或无维度数组
        Err.Clear: On Error GoTo 0
        DictTopN = Array(): Exit Function
    End If
    On Error GoTo 0

    If cnt <= 0 Or n <= 0 Then
        DictTopN = Array()
        Exit Function
    End If

    If n > cnt Then n = cnt
    ReDim result(1 To n, 1 To 2)
    For i = 1 To n
        result(i, 1) = allSorted(i, 1)
        result(i, 2) = allSorted(i, 2)
    Next i
    DictTopN = result
End Function

'=============================================================================
' SetCartesianProduct — 笛卡尔积
'
' 返回: 2D Variant 数组 (1 To n, 1 To 2)，每行是一个配对
'   第 1 列 = A 元素, 第 2 列 = B 元素
'=============================================================================
Public Function SetCartesianProduct( _
    ByRef arrA As Variant, _
    ByRef arrB As Variant) As Variant

    Dim result() As Variant
    Dim i As Long, j As Long
    Dim lbA As Long, ubA As Long, lbB As Long, ubB As Long
    Dim idx As Long

    Dim localA As Variant: localA = arrA
    Dim localB As Variant: localB = arrB
    VK.NormalizeInput localA, True  ' Range→Array + 2D→1D (local copy — §1.1)
    VK.NormalizeInput localB, True
    If Not IsArray(localA) Or Not IsArray(localB) Then
        SetCartesianProduct = Array()
        Exit Function
    End If

    lbA = LBound(localA): ubA = UBound(localA)
    lbB = LBound(localB): ubB = UBound(localB)

    Dim nA As Long, nB As Long
    nA = ubA - lbA + 1: nB = ubB - lbB + 1

    ' 防止 nA * nB Long 溢出 (约 46340² 就会溢出)
    Dim nTotal As Double
    nTotal = CDbl(nA) * CDbl(nB)
    If nTotal <= 0# Or nTotal > 2147483647# Then
        SetCartesianProduct = Array()
        Exit Function
    End If
    ReDim result(1 To CLng(nTotal), 1 To 2)
    idx = 1
    For i = lbA To ubA
        For j = lbB To ubB
            result(idx, 1) = localA(i)
            result(idx, 2) = localB(j)
            idx = idx + 1
        Next j
    Next i
    SetCartesianProduct = result
End Function

'=============================================================================
' 工作表函数 (UDF_DICT_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_DICT_UNION(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    On Error GoTo EH: UDF_DICT_UNION = SetUnion(arr1, arr2): Exit Function
EH: UDF_DICT_UNION = CVErr(xlErrValue)
End Function

Public Function UDF_DICT_INTERSECT(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    On Error GoTo EH: UDF_DICT_INTERSECT = SetIntersect(arr1, arr2): Exit Function
EH: UDF_DICT_INTERSECT = CVErr(xlErrValue)
End Function

Public Function UDF_DICT_DIFFERENCE(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    On Error GoTo EH: UDF_DICT_DIFFERENCE = SetDifference(arr1, arr2): Exit Function
EH: UDF_DICT_DIFFERENCE = CVErr(xlErrValue)
End Function

Public Function UDF_DICT_SYM_DIFF(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    On Error GoTo EH: UDF_DICT_SYM_DIFF = SetSymDifference(arr1, arr2): Exit Function
EH: UDF_DICT_SYM_DIFF = CVErr(xlErrValue)
End Function

Public Function UDF_DICT_CARTESIAN(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    On Error GoTo EH: UDF_DICT_CARTESIAN = SetCartesianProduct(arr1, arr2): Exit Function
EH: UDF_DICT_CARTESIAN = CVErr(xlErrValue)
End Function

Public Function UDF_DICT_ISSUBSET(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    On Error GoTo EH: UDF_DICT_ISSUBSET = SetIsSubset(arr1, arr2): Exit Function
EH: UDF_DICT_ISSUBSET = CVErr(xlErrValue)
End Function

Public Function UDF_DICT_ISEQUAL(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    On Error GoTo EH: UDF_DICT_ISEQUAL = SetEqual(arr1, arr2): Exit Function
EH: UDF_DICT_ISEQUAL = CVErr(xlErrValue)
End Function

Public Function UDF_DICT_GROUPCOUNT(ByVal arr As Variant) As Variant
    On Error GoTo EH: UDF_DICT_GROUPCOUNT = GroupCount(arr): Exit Function
EH: UDF_DICT_GROUPCOUNT = CVErr(xlErrValue)
End Function

'=====================================================================
' 使用示例（期望返回值写在注释中）
'=====================================================================
' Dim d1 As Object, d2 As Object
' Set d1 = DP.Create(): d1.Add "a", 1: d1.Add "b", 2
' Set d2 = DP.Create(): d2.Add "b", 3: d2.Add "c", 4
'
' DictGetDefault(d1, "z", 999)                          → 999
' DictIsEmpty(d1)                                       → False
'
' Dim d3 As Object: Set d3 = DictRenameKey(d1, "a", "x")
' d3("x")                                               → 1
'
' Dim sorted As Variant: sorted = DictSortByValue(d1, False)
' sorted(1, 1), sorted(1, 2)                            → "b", 2
'
' Dim top As Variant: top = DictTopN(d1, 1)
' top(1, 1), top(1, 2)                                  → "b", 2
