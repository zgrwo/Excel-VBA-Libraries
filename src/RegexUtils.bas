Option Explicit

'==============================================================================
' Module:       RegexUtils
' Purpose:      Regular expressions: match, replace, split, capture groups
' Layer:        Text
' Dependencies: VBScript.RegExp (Windows built-in)
' Public:       21 functions/subs
' Notes:        Requires VBScript.RegExp (Windows built-in). No Unicode category support.
'==============================================================================


'=====================================================================
' RegexUtils.bas — VBA 正则表达式工具集
'
' 工作表函数 (UDF_REGEX_*):
'   UDF_REGEX_ISMATCH      — 测试是否存在匹配
'   UDF_REGEX_EXTRACT      — 提取匹配文本
'   UDF_REGEX_EXTRACTALL   — 返回所有匹配数组
'   UDF_REGEX_EXTRACTGROUPS — 提取捕获组 → 2D 数组
'   UDF_REGEX_FULLMATCH    — 测试是否完全匹配
'   UDF_REGEX_REPLACE      — 正则替换
'   UDF_REGEX_SPLIT        — 正则分割
'   UDF_REGEX_COUNT        — 统计匹配次数
'   UDF_REGEX_ESCAPE       — 转义正则特殊字符
'
' 内部函数 (PascalCase，供 VBA 代码调用):
'   RegexIsMatch      — 测试是否存在匹配
'   RegexExtract      — 提取匹配文本
'   RegexExtractAll   — 返回所有匹配数组
'   RegexExtractGroups — 提取捕获组 → 2D 数组
'   RegexIsFullMatch  — 测试是否完全匹配
'   RegexReplace      — 正则替换
'   RegexReplaceRange — 对区域执行替换
'   RegexSplit        — 正则分割
'   RegexSplitToRange — 分割后输出到工作表
'   RegexCount        — 统计匹配次数
'   RegexFindInRange  — 在区域中查找匹配单元格
'   RegexEscape       — 转义正则特殊字符
'
' 依赖: VBScript.RegExp (Windows 系统内置)
'
' 注意: RegexReplaceRange / RegexSplitToRange 不支持多区域 Range (Multi-Area)，
'       传入将导致运行时错误 1004。RegexFindInRange 仅搜索第一个 Area。
'
' 安全警告: VBScript.RegExp 不支持匹配超时。恶意正则模式 (如嵌套量词 "(a+)+b")
' 可能导致灾难性回溯 (ReDoS), 使 Excel 单线程冻结。请勿对不可信来源的
' 正则模式使用本模块函数。
'=====================================================================

Private Const REGEX_PROGID      As String = "VBScript.RegExp"
Private Const ERR_SOURCE        As String = "RegexUtils"
Private Const SEP_DEFAULT       As String = ", "
Private Const ERR_NOT_AVAIL     As Long = vbObjectError + 1001
Private Const ERR_INVALID_REGEX As Long = vbObjectError + 1002
Private Const ERR_INVALID_INPUT  As Long = vbObjectError + 1003
Private Const ERR_REGEX_TOO_LONG As Long = vbObjectError + 1004

' 安全限制: 拒绝过长的正则模式 (防御 ReDoS — VBScript.RegExp 不支持匹配超时)
' 256 字符是保守上限; 实际攻击模式通常在 50-200 字符即可触发指数回溯
Private Const MAX_PATTERN_LENGTH As Long = 256

'=============================================================================
' 辅助函数: 创建并配置正则对象
'=============================================================================
Private Function GetRegex( _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Object  ' 与公共 API 默认值一致

    On Error Resume Next
    Set GetRegex = CreateObject(REGEX_PROGID)
    If Err.Number <> 0 Or GetRegex Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Err.Raise ERR_NOT_AVAIL, ERR_SOURCE, _
            "无法创建 VBScript.RegExp 对象。请确认系统已安装 VBScript 支持组件。"
    End If

    ' 安全防护: 拒绝超长正则模式 (防御 ReDoS — VBScript.RegExp 不支持匹配超时)
    If Len(Pattern) > MAX_PATTERN_LENGTH Then
        On Error GoTo 0
        Err.Raise ERR_REGEX_TOO_LONG, ERR_SOURCE, _
            "正则表达式过长 (" & Len(Pattern) & " > " & MAX_PATTERN_LENGTH & " 字符)。请使用更短的模式。"
    End If

    ' 设置属性 (均在 Resume Next 下，避免非法 Pattern 直接崩溃)
    With GetRegex
        .Pattern = Pattern
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Err.Raise ERR_INVALID_REGEX, ERR_SOURCE, "无效的正则表达式: " & Pattern
        End If
        .IgnoreCase = IgnoreCase
        .MultiLine = MultiLine
        .Global = True
    End With
    On Error GoTo 0

    ' 额外安全测试 (主要检测逻辑错误)
    On Error Resume Next
    GetRegex.Test ""
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        Err.Raise ERR_INVALID_REGEX, ERR_SOURCE, "无效的正则表达式: " & Pattern
    End If
    On Error GoTo 0
End Function

'=============================================================================
' RegexIsMatch — 测试模式是否存在匹配
'=============================================================================
Public Function RegexIsMatch( _
    ByVal InputText As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant

    Dim objReg As Object

    On Error GoTo ErrHandler

    If Len(Pattern) = 0 Then
        RegexIsMatch = False
        Exit Function
    End If
    ' Empty input with valid pattern is valid (e.g., "" matches ".*")

    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
    RegexIsMatch = objReg.Test(InputText)
    Set objReg = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RegexIsMatch", "正则匹配失败: " & Err.Description
End Function

'=============================================================================
' RegexExtract — 正则提取
'
' 参数:
'   InputText - 源文本
'   Pattern   - 正则模式
'   Instance  - 提取第几个匹配:
'                 0 (默认): 返回所有匹配, 用 Separator 连接
'                 >0      : 正向第 N 个 (1-based)
'                 <0      : 倒数第 N 个 (-1 = 最后)
'   Separator - 多结果分隔符 (默认 ", ")
'=============================================================================
Public Function RegexExtract( _
    ByVal InputText As String, _
    ByVal Pattern As String, _
    Optional ByVal Instance As Long = 0, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal Separator As String = SEP_DEFAULT) As Variant

    Dim objReg     As Object
    Dim objMatches As Object
    Dim results()  As String
    Dim i          As Long
    Dim idx        As Long

    On Error GoTo ErrHandler

    If Len(InputText) = 0 Or Len(Pattern) = 0 Then
        RegexExtract = ""
        Exit Function
    End If

    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
    Set objMatches = objReg.Execute(InputText)

    If objMatches.Count = 0 Then
        RegexExtract = ""
        Exit Function
    End If

    If Instance > 0 Then
        If Instance <= objMatches.Count Then
            RegexExtract = objMatches(Instance - 1).Value
        Else
            RegexExtract = ""
        End If

    ElseIf Instance < 0 Then
        idx = objMatches.Count + Instance
        If idx >= 0 Then
            RegexExtract = objMatches(idx).Value
        Else
            RegexExtract = ""
        End If

    Else
        ReDim results(0 To objMatches.Count - 1)
        For i = 0 To objMatches.Count - 1
            results(i) = objMatches(i).Value
        Next i
        RegexExtract = Join(results, Separator)
    End If

    Set objReg = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RegexExtract", "正则提取失败: " & Err.Description
End Function

'=============================================================================
' RegexReplace — 正则替换
'
' 参数:
'   InputText     - 源文本
'   Pattern       - 正则模式
'   Replacement   - 替换文本 (支持 $1, $2... 捕获组反向引用)
'   Instance      - 替换第几个匹配:
'                     0 (默认): 配合 GlobalReplace 决定行为
'                     >0      : 正向第 N 个 (1-based)
'                     <0      : 倒数第 N 个 (-1 = 最后)
'   GlobalReplace - 是否替换全部 (默认 True; Instance <> 0 时忽略)
'
' 已知限制: 当 Instance 指定非零值 (替换单个匹配) 时，若 Pattern 依赖上下文
' (如零宽断言 (?<=...), \b 等) 则可能无法正确执行替换。
'=============================================================================
Public Function RegexReplace( _
    ByVal InputText As String, _
    ByVal Pattern As String, _
    ByVal Replacement As String, _
    Optional ByVal Instance As Long = 0, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal GlobalReplace As Boolean = True) As Variant

    Dim objReg     As Object
    Dim objMatches As Object
    Dim objMatch   As Object
    Dim idx        As Long

    On Error GoTo ErrHandler

    If Len(InputText) = 0 Or Len(Pattern) = 0 Then
        RegexReplace = InputText
        Exit Function
    End If

    ' 快速路径: 全局替换
    If Instance = 0 And GlobalReplace Then
        Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
        RegexReplace = objReg.Replace(InputText, Replacement)
        Exit Function
    End If

    ' Instance <> 0 或仅替换第一个
    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
    objReg.Global = True
    Set objMatches = objReg.Execute(InputText)

    If objMatches.Count = 0 Then
        RegexReplace = InputText
        Exit Function
    End If

    If Instance > 0 Then
        idx = Instance - 1
    ElseIf Instance < 0 Then
        idx = objMatches.Count + Instance
    Else
        idx = 0
    End If

    If idx < 0 Or idx >= objMatches.Count Then
        RegexReplace = InputText
        Exit Function
    End If

    Set objMatch = objMatches(idx)
    With objMatch
        objReg.Global = False
        RegexReplace = Left(InputText, .FirstIndex) & _
                        objReg.Replace(.Value, Replacement) & _
                        Mid(InputText, .FirstIndex + .Length + 1)
        objReg.Global = True
    End With

    Set objReg = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RegexReplace", "正则替换失败: " & Err.Description
End Function

'=============================================================================
' RegexSplit — 正则分割
'
' 返回: Variant 字符串数组 (1-based), 无匹配时返回只含源文本的单元素数组
'=============================================================================
Public Function RegexSplit( _
    ByVal InputText As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant

    Dim objReg   As Object
    Dim result() As String

    On Error GoTo ErrHandler

    If Len(InputText) = 0 Then
        ReDim result(1 To 1)
        result(1) = ""
        RegexSplit = result
        Exit Function
    End If

    If Len(Pattern) = 0 Then
        ReDim result(1 To 1)
        result(1) = InputText
        RegexSplit = result
        Exit Function
    End If

    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
    RegexSplit = RegexSplitWith(InputText, objReg)
    Set objReg = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RegexSplit", "正则分割失败: " & Err.Description
End Function

'=============================================================================
' RegexCount — 统计匹配次数 (支持字符串、Range、数组)
'=============================================================================
Public Function RegexCount( _
    ByVal InputText As Variant, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant

    Dim objReg     As Object
    Dim objMatches As Object
    Dim i As Long, j As Long, total As Long
    Dim v As Variant

    On Error GoTo ErrHandler

    If Len(Pattern) = 0 Then
        RegexCount = 0
        Exit Function
    End If

    ' Range: extract .Value and iterate
    If IsObject(InputText) Then
        If TypeName(InputText) = "Range" Then InputText = InputText.Value
        If Not IsArray(InputText) Then GoTo SingleCell
    End If

    ' Array (from Range.Value or direct array input): iterate and sum
    If IsArray(InputText) Then
        Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
        total = 0
        For i = LBound(InputText, 1) To UBound(InputText, 1)
            For j = LBound(InputText, 2) To UBound(InputText, 2)
                v = InputText(i, j)
                If Not IsError(v) And Not IsNull(v) And Not IsEmpty(v) Then
                    Set objMatches = objReg.Execute(CStr(v))
                    total = total + objMatches.Count
                End If
            Next
        Next
        RegexCount = total
        Set objReg = Nothing
        Exit Function
    End If

SingleCell:
    ' Scalar or single cell
    If IsNull(InputText) Then RegexCount = 0: Exit Function
    If Len(CStr(InputText)) = 0 Then
        RegexCount = 0
        Exit Function
    End If

    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
    Set objMatches = objReg.Execute(CStr(InputText))
    RegexCount = objMatches.Count
    Set objReg = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RegexCount", "正则计数失败: " & Err.Description
End Function

'=============================================================================
' 私有: 使用共享 RegExp 对象执行分割 (避免重复 CreateObject)
'=============================================================================
Private Function RegexSplitWith( _
    ByVal InputText As String, _
    ByVal objReg As Object) As Variant

    Dim objMatches As Object
    Dim objMatch   As Object
    Dim result()   As String
    Dim startPos   As Long
    Dim pieceLen   As Long
    Dim i          As Long
    Dim segCount   As Long

    If Len(InputText) = 0 Then
        ReDim result(1 To 1)
        result(1) = ""
        RegexSplitWith = result
        Exit Function
    End If

    Set objMatches = objReg.Execute(InputText)

    If objMatches.Count = 0 Then
        ReDim result(1 To 1)
        result(1) = InputText
        RegexSplitWith = result
        Exit Function
    End If

    segCount = objMatches.Count + 1
    ReDim result(1 To segCount)
    startPos = 1

    For i = 0 To objMatches.Count - 1
        Set objMatch = objMatches(i)
        pieceLen = objMatch.FirstIndex + 1 - startPos

        If pieceLen > 0 Then
            result(i + 1) = Mid$(InputText, startPos, pieceLen)
        Else
            result(i + 1) = ""
        End If

        startPos = objMatch.FirstIndex + objMatch.Length + 1
    Next i

    If startPos <= Len(InputText) Then
        result(segCount) = Mid$(InputText, startPos)
    Else
        result(segCount) = ""
    End If

    RegexSplitWith = result
End Function

'=============================================================================
' RegexSplitToRange — 将单列文本按正则分割，输出到自动适配的 2D 区域
'
' 用法:
'   RegexSplitToRange Selection, "\|", Range("D1")
'   RegexSplitToRange Range("A1:A10"), ",", Range("B1"), True, True, "N/A"
'
' 参数:
'   srcRange   - 源数据区域 (单列，含空单元格会自动跳过)
'   Pattern    - 正则分割模式
'   destCell   - 输出起始单元格
'   IgnoreCase - 忽略大小写 (默认 True)
'   MultiLine  - 多行模式 (默认 True)
'   fillValue  - 不足最大列数的单元格填充值 (默认 "")
'
' 输出: 从 destCell 开始自动适配大小的区域，第1行为分割结果
'=============================================================================
Public Sub RegexSplitToRange( _
    ByVal srcRange As Range, _
    ByVal Pattern As String, _
    ByVal destCell As Range, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal fillValue As String = "")

    Dim srcData As Variant
    Dim nRows As Long, i As Long
    Dim objReg As Object
    Dim allParts() As Variant
    Dim maxCols As Long
    Dim cellText As String
    Dim parts As Variant
    Dim temp() As String
    Dim j As Long, nParts As Long
    Dim outArr() As Variant
    Dim rowArr As Variant
    Dim lbRow As Long

    On Error GoTo ErrHandler

    If srcRange Is Nothing Or destCell Is Nothing Then Exit Sub
    If Len(Pattern) = 0 Then Exit Sub
    If srcRange.Columns.Count > 1 Then
        Err.Raise ERR_INVALID_INPUT, "RegexSplitToRange", _
            "源区域必须为单列，当前为 " & srcRange.Columns.Count & " 列。"
    End If

    If srcRange.Cells.Count = 1 Then
        ReDim srcData(1 To 1, 1 To 1)
        srcData(1, 1) = srcRange.Value
        nRows = 1
    Else
        srcData = srcRange.Value
        nRows = UBound(srcData, 1)
    End If

    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)

    ReDim allParts(1 To nRows)
    maxCols = 0

    For i = 1 To nRows
        cellText = ""
        If Not IsEmpty(srcData(i, 1)) And Not IsError(srcData(i, 1)) Then
            cellText = CStr(srcData(i, 1))
        End If

        If Len(cellText) = 0 Then
            ReDim temp(1 To 1)
            temp(1) = ""
            allParts(i) = temp
            If 1 > maxCols Then maxCols = 1
        Else
            parts = RegexSplitWith(cellText, objReg)
            allParts(i) = parts
            nParts = UBound(parts) - LBound(parts) + 1
            If nParts > maxCols Then maxCols = nParts
        End If
    Next i

    If maxCols = 0 Then maxCols = 1

    ReDim outArr(1 To nRows, 1 To maxCols)

    For i = 1 To nRows
        rowArr = allParts(i)
        nParts = UBound(rowArr) - LBound(rowArr) + 1
        lbRow = LBound(rowArr)
        For j = 1 To maxCols
            If j <= nParts Then
                outArr(i, j) = rowArr(lbRow + j - 1)
            Else
                outArr(i, j) = fillValue
            End If
        Next j
    Next i

    destCell.Resize(nRows, maxCols).Value = outArr
    Set objReg = Nothing
    Exit Sub

ErrHandler:
    If Not objReg Is Nothing Then Set objReg = Nothing
    Err.Raise Err.Number, "RegexSplitToRange", "正则分割到区域失败: " & Err.Description
End Sub

'=============================================================================
' RegexExtractAll — 返回所有匹配的数组 (不拼接)
'=============================================================================
Public Function RegexExtractAll( _
    ByVal InputText As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant

    Dim objReg As Object, objMatches As Object
    Dim result() As String
    Dim i As Long

    On Error GoTo ErrHandler

    If Len(InputText) = 0 Or Len(Pattern) = 0 Then
        RegexExtractAll = Array()
        Exit Function
    End If

    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
    Set objMatches = objReg.Execute(InputText)

    If objMatches.Count = 0 Then
        RegexExtractAll = Array()
    Else
        ReDim result(0 To objMatches.Count - 1)
        For i = 0 To objMatches.Count - 1
            result(i) = objMatches(i).Value
        Next i
        RegexExtractAll = result
    End If
    Set objReg = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RegexExtractAll", "正则提取失败: " & Err.Description
End Function

'=============================================================================
' RegexExtractGroups — 提取捕获组 (SubMatches) 为 2D 数组
'
' 返回: (0 To 匹配数-1, 0 To 最大捕获组数-1) 的 2D Variant 数组
'       每行为一次匹配的捕获组内容 (SubMatches)。
'       注意: 不含完整匹配文本；当无捕获组时返回空数组 Array()。
'=============================================================================
Public Function RegexExtractGroups( _
    ByVal InputText As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant

    Dim objReg As Object, objMatches As Object, objMatch As Object
    Dim result() As Variant
    Dim i As Long, j As Long, nMatches As Long, maxGroups As Long

    On Error GoTo ErrHandler

    If Len(InputText) = 0 Or Len(Pattern) = 0 Then
        RegexExtractGroups = Array()
        Exit Function
    End If

    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
    Set objMatches = objReg.Execute(InputText)

    nMatches = objMatches.Count
    If nMatches = 0 Then
        RegexExtractGroups = Array()
        Exit Function
    End If

    ' 找出最大的捕获组数
    For i = 0 To nMatches - 1
        With objMatches(i)
            If .SubMatches.Count > maxGroups Then
                maxGroups = .SubMatches.Count
            End If
        End With
    Next i

    If maxGroups = 0 Then
        RegexExtractGroups = Array(): Exit Function
    End If

    ReDim result(0 To nMatches - 1, 0 To maxGroups - 1)

    For i = 0 To nMatches - 1
        Set objMatch = objMatches(i)
        With objMatch
            For j = 0 To .SubMatches.Count - 1
                result(i, j) = .SubMatches(j)
            Next j
        End With
    Next i
    RegexExtractGroups = result
    Set objReg = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RegexExtractGroups", "正则分组提取失败: " & Err.Description
End Function

'=============================================================================
' RegexIsFullMatch — 测试整个字符串是否完全匹配模式
'=============================================================================
Public Function RegexIsFullMatch( _
    ByVal InputText As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant

    Dim objReg As Object, objMatches As Object

    If Len(InputText) = 0 And Len(Pattern) = 0 Then
        RegexIsFullMatch = True
        Exit Function
    End If

    On Error GoTo ErrHandler
    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
    ' 关闭 Global 以避免可空模式 (.*) 产生零宽尾匹配 (#77)
    objReg.Global = False
    Set objMatches = objReg.Execute(InputText)

    If objMatches.Count = 1 Then
        RegexIsFullMatch = (objMatches(0).FirstIndex = 0 And _
                             objMatches(0).Length = Len(InputText))
    ElseIf objMatches.Count = 0 Then
        RegexIsFullMatch = (Len(InputText) = 0 And Len(Pattern) = 0)
    Else
        ' 多个匹配 — 检查是否整个输入被一个匹配覆盖
        ' (Global=False 不应出现, 但安全处理)
        RegexIsFullMatch = False
    End If
    Set objReg = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RegexIsFullMatch", "正则匹配失败: " & Err.Description
End Function

'=============================================================================
' RegexEscape — 转义正则表达式特殊字符
'=============================================================================
Public Function RegexEscape(ByVal text As String) As Variant
    Const SPECIAL As String = "\.^$*+?{}[]()|"
    Dim i As Long
    Dim ch As String
    Dim parts() As String
    Dim pIdx As Long

    If Len(text) = 0 Then
        RegexEscape = ""
        Exit Function
    End If

    On Error GoTo ErrHandler

    pIdx = 0
    ReDim parts(0 To Len(text) * 2)
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If InStr(1, SPECIAL, ch, vbBinaryCompare) > 0 Then
            parts(pIdx) = "\": pIdx = pIdx + 1
        End If
        parts(pIdx) = ch: pIdx = pIdx + 1
    Next i
    If pIdx > 0 Then
        ReDim Preserve parts(0 To pIdx - 1)
        RegexEscape = Join(parts, "")
    End If

    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RegexEscape", "正则转义失败: " & Err.Description
End Function

'=============================================================================
' RegexFindInRange — 在区域中查找匹配正则的单元格
'
' 返回: 匹配单元格的 Range，未找到返回 Nothing
' 限制: rng.Areas.Count > 1 时将引发运行时错误 1004，仅支持单区域 Range。
'=============================================================================
Public Function RegexFindInRange( _
    ByRef rng As Range, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Range

    Dim objReg As Object
    Dim addrParts() As String
    Dim apIdx As Long
    Dim batchResult As Range
    Dim batchAddr As String
    Dim batchStart As Long, batchEnd As Long
    Dim bIdx As Long
    Dim batchParts() As String
    Const BATCH_SIZE As Long = 30

    On Error GoTo ErrHandler

    If rng Is Nothing Or Len(Pattern) = 0 Then
        Set RegexFindInRange = Nothing
        Exit Function
    End If

    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)

    Dim data As Variant
    Dim nR As Long, nC As Long
    If rng.Count = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = rng.Value
        nR = 1: nC = 1
    Else
        data = rng.Value
        nR = UBound(data, 1): nC = UBound(data, 2)
    End If

    apIdx = 0
    ReDim addrParts(0 To nR * nC)
    Dim iR As Long, iC As Long
    For iR = 1 To nR
        For iC = 1 To nC
            If Not IsEmpty(data(iR, iC)) Then
                If Not IsError(data(iR, iC)) Then
                    If objReg.Test(CStr(data(iR, iC))) Then
                        addrParts(apIdx) = rng.Cells(iR, iC).Address
                        apIdx = apIdx + 1
                    End If
                End If
            End If
        Next iC
    Next iR

    If apIdx > 0 Then
        ReDim Preserve addrParts(0 To apIdx - 1)
        Dim ws As Worksheet: Set ws = rng.Parent
        If apIdx <= BATCH_SIZE Then
            Set RegexFindInRange = ws.Range(Join(addrParts, ","))
        Else
            Set batchResult = Nothing
            For batchStart = 0 To apIdx - 1 Step BATCH_SIZE
                batchEnd = batchStart + BATCH_SIZE - 1
                If batchEnd >= apIdx Then batchEnd = apIdx - 1
                ReDim batchParts(0 To batchEnd - batchStart)
                For bIdx = batchStart To batchEnd
                    batchParts(bIdx - batchStart) = addrParts(bIdx)
                Next bIdx
                batchAddr = Join(batchParts, ",")
                If batchResult Is Nothing Then
                    Set batchResult = ws.Range(batchAddr)
                Else
                    Set batchResult = Application.Union(batchResult, ws.Range(batchAddr))
                End If
            Next batchStart
            Set RegexFindInRange = batchResult
        End If
    Else
        Set RegexFindInRange = Nothing
    End If
    Set objReg = Nothing
    Exit Function

ErrHandler:
    Set RegexFindInRange = Nothing
End Function

'=============================================================================
' RegexReplaceRange — 对区域中每个单元格执行正则替换
'
' 参数:
'   rng         - 目标区域 (就地修改)
'   Pattern     - 正则模式
'   Replacement - 替换文本
'
' 限制: 不支持多区域 Range (Multi-Area)，rng.Areas.Count > 1 时将引发
'       运行时错误 1004。
'=============================================================================
Public Sub RegexReplaceRange( _
    ByRef rng As Range, _
    ByVal Pattern As String, _
    ByVal Replacement As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True)

    Dim objReg As Object
    Dim data As Variant
    Dim i As Long, j As Long
    Dim nRows As Long, nCols As Long
    Dim hasAnyFormula As Boolean

    On Error GoTo ErrHandler

    If rng Is Nothing Or Len(Pattern) = 0 Then Exit Sub

    Set objReg = GetRegex(Pattern, IgnoreCase, MultiLine)
    objReg.Global = True

    If rng.Count = 1 Then
        If Not rng.HasFormula Then
            rng.Value = objReg.Replace(CStr(rng.Value), Replacement)
        End If
    Else
        data = rng.Value
        nRows = UBound(data, 1): nCols = UBound(data, 2)
        ' rng.HasFormula 在混合内容时返回 Null — 赋值给 Boolean 会触发 Error 94
        Dim hasFormulaVar As Variant: hasFormulaVar = rng.HasFormula
        If Not IsNull(hasFormulaVar) Then hasAnyFormula = CBool(hasFormulaVar)

        For i = 1 To nRows
            For j = 1 To nCols
                If Not IsEmpty(data(i, j)) And Not IsError(data(i, j)) Then
                    If Not hasAnyFormula Then
                        data(i, j) = objReg.Replace(CStr(data(i, j)), Replacement)
                    Else
                        With rng.Cells(i, j)
                            If Not .HasFormula Then data(i, j) = objReg.Replace(CStr(data(i, j)), Replacement)
                        End With
                    End If
                End If
            Next j
        Next i

        If Not hasAnyFormula Then
            rng.Value = data
        Else
            For i = 1 To nRows
                For j = 1 To nCols
                    With rng.Cells(i, j)
                        If Not .HasFormula Then .Value = data(i, j)
                    End With
                Next j
            Next i
        End If
    End If
    Set objReg = Nothing
    Exit Sub

ErrHandler:
    Dim errNum As Long:    errNum = Err.Number
    Dim errSrc As String:  errSrc = Err.Source
    Dim errDesc As String: errDesc = Err.Description
    If Not objReg Is Nothing Then Set objReg = Nothing
    Err.Raise errNum, errSrc, errDesc
End Sub

'=============================================================================
' 工作表函数 (UDF_REGEX_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_REGEX_ISMATCH( _
    ByVal InputText As Variant, _
    ByVal Pattern As Variant, _
    Optional ByVal IgnoreCase As Variant = True, _
    Optional ByVal MultiLine As Variant = True) As Variant
    On Error GoTo EH
    If IsObject(InputText) Then If TypeName(InputText) = "Range" Then InputText = InputText.Value
    If IsArray(InputText) Then
        Dim i As Long, j As Long, v As Variant, resultArr() As Variant
        ReDim resultArr(LBound(InputText,1) To UBound(InputText,1), LBound(InputText,2) To UBound(InputText,2))
        For i = LBound(InputText,1) To UBound(InputText,1)
            For j = LBound(InputText,2) To UBound(InputText,2)
                v = InputText(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RegexIsMatch(CStr(v), Pattern, IgnoreCase, MultiLine)
        Next: Next
        UDF_REGEX_ISMATCH = resultArr: Exit Function
    End If
    UDF_REGEX_ISMATCH = RegexIsMatch(CStr(InputText), Pattern, IgnoreCase, MultiLine): Exit Function
EH: UDF_REGEX_ISMATCH = CVErr(xlErrValue)
End Function

Public Function UDF_REGEX_EXTRACT( _
    ByVal InputText As Variant, _
    ByVal Pattern As Variant, _
    Optional ByVal Instance As Variant = 0, _
    Optional ByVal IgnoreCase As Variant = True, _
    Optional ByVal MultiLine As Variant = True, _
    Optional ByVal Separator As Variant = SEP_DEFAULT) As Variant
    On Error GoTo EH
    If IsObject(InputText) Then If TypeName(InputText) = "Range" Then InputText = InputText.Value
    If IsArray(InputText) Then
        Dim i As Long, j As Long, v As Variant, resultArr() As Variant
        ReDim resultArr(LBound(InputText,1) To UBound(InputText,1), LBound(InputText,2) To UBound(InputText,2))
        For i = LBound(InputText,1) To UBound(InputText,1)
            For j = LBound(InputText,2) To UBound(InputText,2)
                v = InputText(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RegexExtract(CStr(v), Pattern, Instance, IgnoreCase, MultiLine, Separator)
        Next: Next
        UDF_REGEX_EXTRACT = resultArr: Exit Function
    End If
    UDF_REGEX_EXTRACT = RegexExtract(CStr(InputText), Pattern, Instance, IgnoreCase, MultiLine, Separator): Exit Function
EH: UDF_REGEX_EXTRACT = CVErr(xlErrValue)
End Function

Public Function UDF_REGEX_EXTRACTALL( _
    ByVal InputText As Variant, _
    ByVal Pattern As Variant, _
    Optional ByVal IgnoreCase As Variant = True, _
    Optional ByVal MultiLine As Variant = True) As Variant
    On Error GoTo EH
    If IsObject(InputText) Then If TypeName(InputText) = "Range" Then InputText = InputText.Value
    If IsArray(InputText) Then
        Dim i As Long, j As Long, v As Variant, resultArr() As Variant
        ReDim resultArr(LBound(InputText,1) To UBound(InputText,1), LBound(InputText,2) To UBound(InputText,2))
        For i = LBound(InputText,1) To UBound(InputText,1)
            For j = LBound(InputText,2) To UBound(InputText,2)
                v = InputText(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RegexExtractAll(CStr(v), Pattern, IgnoreCase, MultiLine)
        Next: Next
        UDF_REGEX_EXTRACTALL = resultArr: Exit Function
    End If
    UDF_REGEX_EXTRACTALL = RegexExtractAll(CStr(InputText), Pattern, IgnoreCase, MultiLine): Exit Function
EH: UDF_REGEX_EXTRACTALL = CVErr(xlErrValue)
End Function

Public Function UDF_REGEX_EXTRACTGROUPS( _
    ByVal InputText As Variant, _
    ByVal Pattern As Variant, _
    Optional ByVal IgnoreCase As Variant = True, _
    Optional ByVal MultiLine As Variant = True) As Variant
    On Error GoTo EH
    If IsObject(InputText) Then If TypeName(InputText) = "Range" Then InputText = InputText.Value
    If IsArray(InputText) Then
        Dim i As Long, j As Long, v As Variant, resultArr() As Variant
        ReDim resultArr(LBound(InputText,1) To UBound(InputText,1), LBound(InputText,2) To UBound(InputText,2))
        For i = LBound(InputText,1) To UBound(InputText,1)
            For j = LBound(InputText,2) To UBound(InputText,2)
                v = InputText(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RegexExtractGroups(CStr(v), Pattern, IgnoreCase, MultiLine)
        Next: Next
        UDF_REGEX_EXTRACTGROUPS = resultArr: Exit Function
    End If
    UDF_REGEX_EXTRACTGROUPS = RegexExtractGroups(CStr(InputText), Pattern, IgnoreCase, MultiLine): Exit Function
EH: UDF_REGEX_EXTRACTGROUPS = CVErr(xlErrValue)
End Function

Public Function UDF_REGEX_FULLMATCH( _
    ByVal InputText As Variant, _
    ByVal Pattern As Variant, _
    Optional ByVal IgnoreCase As Variant = True, _
    Optional ByVal MultiLine As Variant = True) As Variant
    On Error GoTo EH
    If IsObject(InputText) Then If TypeName(InputText) = "Range" Then InputText = InputText.Value
    If IsArray(InputText) Then
        Dim i As Long, j As Long, v As Variant, resultArr() As Variant
        ReDim resultArr(LBound(InputText,1) To UBound(InputText,1), LBound(InputText,2) To UBound(InputText,2))
        For i = LBound(InputText,1) To UBound(InputText,1)
            For j = LBound(InputText,2) To UBound(InputText,2)
                v = InputText(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RegexIsFullMatch(CStr(v), Pattern, IgnoreCase, MultiLine)
        Next: Next
        UDF_REGEX_FULLMATCH = resultArr: Exit Function
    End If
    UDF_REGEX_FULLMATCH = RegexIsFullMatch(CStr(InputText), Pattern, IgnoreCase, MultiLine): Exit Function
EH: UDF_REGEX_FULLMATCH = CVErr(xlErrValue)
End Function

Public Function UDF_REGEX_REPLACE( _
    ByVal InputText As Variant, _
    ByVal Pattern As Variant, _
    ByVal Replacement As Variant, _
    Optional ByVal Instance As Variant = 0, _
    Optional ByVal IgnoreCase As Variant = True, _
    Optional ByVal MultiLine As Variant = True, _
    Optional ByVal GlobalReplace As Variant = True) As Variant
    On Error GoTo EH
    If IsObject(InputText) Then If TypeName(InputText) = "Range" Then InputText = InputText.Value
    If IsArray(InputText) Then
        Dim i As Long, j As Long, v As Variant, resultArr() As Variant
        ReDim resultArr(LBound(InputText,1) To UBound(InputText,1), LBound(InputText,2) To UBound(InputText,2))
        For i = LBound(InputText,1) To UBound(InputText,1)
            For j = LBound(InputText,2) To UBound(InputText,2)
                v = InputText(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RegexReplace(CStr(v), Pattern, Replacement, Instance, IgnoreCase, MultiLine, GlobalReplace)
        Next: Next
        UDF_REGEX_REPLACE = resultArr: Exit Function
    End If
    UDF_REGEX_REPLACE = RegexReplace(CStr(InputText), Pattern, Replacement, Instance, IgnoreCase, MultiLine, GlobalReplace): Exit Function
EH: UDF_REGEX_REPLACE = CVErr(xlErrValue)
End Function

Public Function UDF_REGEX_SPLIT( _
    ByVal InputText As Variant, _
    ByVal Pattern As Variant, _
    Optional ByVal IgnoreCase As Variant = True, _
    Optional ByVal MultiLine As Variant = True) As Variant
    On Error GoTo EH
    If IsObject(InputText) Then If TypeName(InputText) = "Range" Then InputText = InputText.Value
    If IsArray(InputText) Then
        Dim i As Long, j As Long, v As Variant, resultArr() As Variant
        ReDim resultArr(LBound(InputText,1) To UBound(InputText,1), LBound(InputText,2) To UBound(InputText,2))
        For i = LBound(InputText,1) To UBound(InputText,1)
            For j = LBound(InputText,2) To UBound(InputText,2)
                v = InputText(i,j): If IsError(v) Or IsNull(v) Or IsEmpty(v) Then resultArr(i,j) = v Else resultArr(i,j) = RegexSplit(CStr(v), Pattern, IgnoreCase, MultiLine)
        Next: Next
        UDF_REGEX_SPLIT = resultArr: Exit Function
    End If
    UDF_REGEX_SPLIT = RegexSplit(CStr(InputText), Pattern, IgnoreCase, MultiLine): Exit Function
EH: UDF_REGEX_SPLIT = CVErr(xlErrValue)
End Function

Public Function UDF_REGEX_COUNT( _
    ByVal InputText As Variant, _
    ByVal Pattern As Variant, _
    Optional ByVal IgnoreCase As Variant = True, _
    Optional ByVal MultiLine As Variant = True) As Variant
    On Error GoTo EH: UDF_REGEX_COUNT = RegexCount(InputText, Pattern, IgnoreCase, MultiLine): Exit Function
EH: UDF_REGEX_COUNT = CVErr(xlErrValue)
End Function

Public Function UDF_REGEX_ESCAPE(ByVal text As Variant) As Variant
    On Error GoTo EH: UDF_REGEX_ESCAPE = RegexEscape(text): Exit Function
EH: UDF_REGEX_ESCAPE = CVErr(xlErrValue)
End Function