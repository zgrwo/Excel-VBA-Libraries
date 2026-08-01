Option Explicit

'==============================================================================
' Module:       JsonUtils
' Purpose:      JSON: pure-VBA recursive descent parser and stringifier
' Layer:        Text
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       10 functions/subs
'==============================================================================

Private DP As New DictProxy

'=====================================================================
' JsonUtils.bas — JSON 解析 (纯 VBA 递归下降解析器 · 重构优化版)
'
' 纯 VBA，无外部依赖。支持对象/数组/字符串/数字/布尔/null/嵌套。
' 支持 UTF-16 代理对 (Emoji / CJK Ext-B 等补充平面字符)。
'
' 【关键设计】
'   1. 状态封装 — UDT 传递解析状态，消除模块级变量，线程/递归安全
'   2. VarLetSet — 自动判 Set/Let，仅限局部变量；函数返回内联 IsObject
'   3. 区域安全 — Val() 解析数字，德语/法语等逗号小数点系统不崩溃
'   4. 严格验证 — 数字格式校验 (禁前导零/必须含小数位/指数位)
'   5. 完整 Unicode — 支持代理对 (Surrogate Pairs)，正确解析 Emoji/生僻字
'   6. 快速路径 — 无转义字符串直接 Mid$ 截取，无 Join 拼接开销
'   7. 重复键覆盖 — dict.Add 替代 dict(key)=val，兼容对象值
'
' 工作表函数 (UDF_JSON_*):
'   UDF_JSON_GET       — 按路径提取值 (对象/数组返回占位文本)
'   UDF_JSON_KEYS      — 列出对象所有键
'   UDF_JSON_IS_VALID  — 校验合法性
'
' VBA-only (Public, 返回 Dictionary/Object):
'   JsonParse   — 解析 → Variant (Dictionary / Array / 标量)
'   JsonGet     — 路径提取 (支持 a.b / arr[0] / users[0].name)
'   JsonToRange — JSON 对象数组 → 工作表
'   JsonGetKeys    — 获取键列表 As String()
'   JsonIsValid    — 布尔校验 (不抛错)
'   JsonStringify  — VBA Variant → JSON 字符串 (VBA-only)
'=====================================================================

' --- 错误码 ---
Private Const ERR_INVALID_JSON  As Long = vbObjectError + 1301
Private Const ERR_PATH_NOT_FOUND As Long = vbObjectError + 1302
Private Const ERR_INVALID_PATH  As Long = vbObjectError + 1303
Private Const ERR_FORMULA_RANGE As Long = vbObjectError + 1304
Private Const ERR_INVALID_INPUT As Long = vbObjectError + 1001

' --- 递归深度保护 ---
Private Const MAX_JSON_DEPTH As Long = 512

' --- 解析状态 (UDT，全部状态封装在 UDT 中，消除模块级变量) ---
Private Type TJsonState
    jsonText As String
    pos      As Long
    textLen  As Long
    depth    As Long  ' 当前递归深度（替代原 mParseDepth 模块级变量）
End Type

'=====================================================================
' Private 辅助函数

' VarLetSet — 安全赋值 Variant (Set for objects, Let otherwise)
' Source is ByVal to avoid "array locked" error from temporary Variants.
Private Sub VarLetSet(ByRef target As Variant, ByVal source As Variant)
    If IsObject(source) Then
        Set target = source
    Else
        target = source
    End If
End Sub

' --- 流操作 (基于 TJsonState) ---

Private Sub SkipWhitespace(ByRef st As TJsonState)
    Do While st.pos <= st.textLen
        Select Case AscW(Mid$(st.jsonText, st.pos, 1))
            Case 9, 10, 13, 32: st.pos = st.pos + 1
            Case Else: Exit Do
        End Select
    Loop
End Sub

Private Function PeekChar(ByRef st As TJsonState) As String
    If st.pos <= st.textLen Then PeekChar = Mid$(st.jsonText, st.pos, 1) Else PeekChar = ""
End Function

Private Function ReadChar(ByRef st As TJsonState) As String
    ReadChar = PeekChar(st)
    If Len(ReadChar) > 0 Then st.pos = st.pos + 1
End Function

Private Sub ExpectChar(ByRef st As TJsonState, ByVal expected As String)
    If ReadChar(st) <> expected Then
        Err.Raise ERR_INVALID_JSON, "JsonUtils", _
            "JSON 语法错误: 位置 " & (st.pos - 1) & " 期望 " & Chr$(39) & expected & Chr$(39)
    End If
End Sub

'=====================================================================
' 核心解析器 (全部接收 ByRef st)
'=====================================================================

' ParseValue — 值分发器 (ByRef outVar 消除双重解析 Bug)
Private Sub ParseValue(ByRef st As TJsonState, ByRef outVar As Variant)
    Dim c As String

    SkipWhitespace st
    c = PeekChar(st)
    Select Case c
        Case "{"
            Set outVar = ParseObject(st)
        Case "["
            outVar = ParseArray(st)
        Case """"
            outVar = ParseString(st)
        Case "t", "f"
            outVar = ParseLiteral(st)
        Case "n"
            outVar = ParseNull(st)
        Case Else
            If c = "-" Or (c >= "0" And c <= "9") Then
                outVar = ParseNumber(st)
            Else
                Err.Raise ERR_INVALID_JSON, "JsonUtils", _
                    "JSON 语法错误: 位置 " & st.pos & " 意外字符 " & Chr$(39) & c & Chr$(39)
            End If
    End Select
    SkipWhitespace st
End Sub

' ParseObject — 解析 JSON 对象 → Dictionary (重复键自动覆盖)
Private Function ParseObject(ByRef st As TJsonState) As Object
    Dim dict As Object
    Dim key As String
    Dim val As Variant

    st.depth = st.depth + 1
    If st.depth > MAX_JSON_DEPTH Then
        st.depth = st.depth - 1
        Err.Raise ERR_INVALID_JSON, "JsonUtils", _
            "JSON 嵌套深度超过限制 (" & MAX_JSON_DEPTH & ")"
    End If

    Set dict = DP.Create()
    ExpectChar st, "{"
    SkipWhitespace st
    If PeekChar(st) = "}" Then
        ReadChar st
        Set ParseObject = dict
        st.depth = st.depth - 1
        Exit Function
    End If

    Do
        SkipWhitespace st
        key = ParseString(st)
        SkipWhitespace st
        ExpectChar st, ":"
        ParseValue st, val
        ' Use Add/Remove — dict.Add handles Variant objects correctly
        ' (Let-assignment dict(key)=val triggers default property for objects)
        If dict.Exists(key) Then dict.Remove key
        dict.Add key, val       ' 已存在则覆盖 (容错)
        SkipWhitespace st
        If PeekChar(st) = "}" Then
            ReadChar st
            Exit Do
        End If
        ExpectChar st, ","
    Loop
    Set ParseObject = dict
    st.depth = st.depth - 1
End Function

' ParseArray — 解析 JSON 数组 → Variant Array
Private Function ParseArray(ByRef st As TJsonState) As Variant
    Dim items() As Variant
    Dim emptyArr() As Variant
    Dim cnt As Long
    Dim cap As Long

    st.depth = st.depth + 1
    If st.depth > MAX_JSON_DEPTH Then
        st.depth = st.depth - 1
        Err.Raise ERR_INVALID_JSON, "JsonUtils", _
            "JSON 嵌套深度超过限制 (" & MAX_JSON_DEPTH & ")"
    End If

    cap = 16
    cnt = 0
    ReDim items(0 To cap - 1)
    ExpectChar st, "["
    SkipWhitespace st
    If PeekChar(st) = "]" Then
        ReadChar st
        ' Return uninitialized empty array — safer than Split(vbNullString)
        ' (Split(vbNullString) relies on undocumented behavior, LBound/UBound may vary by host)
        ParseArray = emptyArr
        st.depth = st.depth - 1
        Exit Function
    End If

    Do
        SkipWhitespace st
        If cnt >= cap Then
            cap = cap * 2
            ReDim Preserve items(0 To cap - 1)
        End If
        ParseValue st, items(cnt)
        cnt = cnt + 1
        SkipWhitespace st
        If PeekChar(st) = "]" Then
            ReadChar st
            Exit Do
        End If
        ExpectChar st, ","
    Loop

    If cnt > 0 Then
        ReDim Preserve items(0 To cnt - 1)
        ParseArray = items
    Else
        ParseArray = emptyArr
    End If
    st.depth = st.depth - 1
End Function

' ParseString — 解析 JSON 字符串 (快速路径 + 完整 Unicode)
Private Function ParseString(ByRef st As TJsonState) As String
    Dim startPos As Long
    Dim parts()   As String
    Dim pIdx      As Long
    Dim c         As String
    Dim cp        As Long
    Dim hexStr    As String
    Dim hexStr2   As String
    Dim cp2       As Long

    ExpectChar st, """"
    startPos = st.pos

    ' === 快速路径：扫描未转义的结束引号 ===
    Do While st.pos <= st.textLen
        Select Case AscW(Mid$(st.jsonText, st.pos, 1))
            Case 34 ' "
                ParseString = Mid$(st.jsonText, startPos, st.pos - startPos)
                st.pos = st.pos + 1
                Exit Function
            Case 92 ' \
                Exit Do
        End Select
        st.pos = st.pos + 1
    Loop

    If st.pos > st.textLen Then
        Err.Raise ERR_INVALID_JSON, "JsonUtils", "JSON 语法错误: 字符串未闭合"
    End If

    ' === 慢速路径：处理含转义字符的字符串 ===
    ' Intentionally oversized to avoid ReDim Preserve in the escape-sequence loop;
    ' final trim at line ~263 handles the overallocation
    ReDim parts(0 To st.textLen - startPos)
    pIdx = 0

    ' 保留快速路径已扫描的无转义部分
    If st.pos > startPos Then
        parts(0) = Mid$(st.jsonText, startPos, st.pos - startPos)
        pIdx = 1
    End If

    Do While st.pos <= st.textLen
        c = ReadChar(st)
        If c = """" Then
            If pIdx > 0 Then ReDim Preserve parts(0 To pIdx - 1)
            ParseString = Join(parts, "")
            Exit Function
        End If
        If c = "\" Then
            c = ReadChar(st)
            Select Case c
                Case """": parts(pIdx) = """":    pIdx = pIdx + 1
                Case "\":  parts(pIdx) = "\":     pIdx = pIdx + 1
                Case "/":  parts(pIdx) = "/":     pIdx = pIdx + 1
                Case "b":  parts(pIdx) = Chr$(8): pIdx = pIdx + 1
                Case "f":  parts(pIdx) = Chr$(12): pIdx = pIdx + 1
                Case "n":  parts(pIdx) = vbLf:    pIdx = pIdx + 1
                Case "r":  parts(pIdx) = vbCr:    pIdx = pIdx + 1
                Case "t":  parts(pIdx) = vbTab:   pIdx = pIdx + 1
                Case "u":
                    If st.pos + 3 > st.textLen Then
                        Err.Raise ERR_INVALID_JSON, "JsonUtils", _
                            "JSON 语法错误: 不完整的 Unicode 转义"
                    End If
                    hexStr = Mid$(st.jsonText, st.pos, 4)
                    If Not (hexStr Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]") Then
                        Err.Raise ERR_INVALID_JSON, "JsonUtils", _
                            "JSON 语法错误: 无效的 Unicode 转义 \\u" & hexStr
                    End If
                    cp = CLng("&H" & hexStr)
                    st.pos = st.pos + 4

                    ' 处理 UTF-16 代理对 (Emoji / 生僻字)
                    ' 组合后的码位 > U+FFFF 需要拆回代理对写入 VBA 字符串
                    Dim isSurrogatePair As Boolean
                    isSurrogatePair = False
                    If cp >= &HD800 And cp <= &HDBFF Then
                        If Mid$(st.jsonText, st.pos, 2) = "\u" Then
                            hexStr2 = Mid$(st.jsonText, st.pos + 2, 4)
                            If hexStr2 Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                                cp2 = CLng("&H" & hexStr2)
                                If cp2 >= &HDC00 And cp2 <= &HDFFF Then
                                    cp = &H10000 + ((cp - &HD800) * &H400) + (cp2 - &HDC00)
                                    st.pos = st.pos + 6
                                    isSurrogatePair = True
                                End If
                            End If
                        End If
                    End If

                    If isSurrogatePair Then
                        ' 补充平面字符: 拆为 UTF-16 高低代理项
                        Dim spAdj As Long: spAdj = cp - &H10000
                        parts(pIdx) = ChrW$(&HD800& + (spAdj \ &H400&)): pIdx = pIdx + 1
                        parts(pIdx) = ChrW$(&HDC00& + (spAdj And &H3FF&)): pIdx = pIdx + 1
                    Else
                        parts(pIdx) = ChrW$(cp)
                        pIdx = pIdx + 1
                    End If
                Case Else
                    Err.Raise ERR_INVALID_JSON, "JsonUtils", _
                        "JSON 语法错误: 位置 " & (st.pos - 1) & " 无效的转义字符 \" & c
            End Select
        Else
            parts(pIdx) = c
            pIdx = pIdx + 1
        End If
    Loop
    Err.Raise ERR_INVALID_JSON, "JsonUtils", "JSON 语法错误: 字符串未闭合"
End Function

' ParseNumber — 解析 JSON 数字 (严格格式校验 + Val() 区域安全)
Private Function ParseNumber(ByRef st As TJsonState) As Variant
    Dim start    As Long
    Dim hasDigit As Boolean
    Dim hasFrac  As Boolean
    Dim hasExp   As Boolean
    Dim d        As Long
    Dim numStr   As String

    start = st.pos

    ' 可选负号
    If PeekChar(st) = "-" Then ReadChar st
    If PeekChar(st) = "+" Then Err.Raise ERR_INVALID_JSON, "JsonUtils", _
        "JSON 数字不允许前导 ""+"" 号 (位置 " & st.pos & ")"

    ' 整数部分 (禁止前导零，除非整体就是 0)
    hasDigit = False
    If PeekChar(st) = "0" Then
        ReadChar st
        hasDigit = True
    Else
        Do While st.pos <= st.textLen
            d = AscW(Mid$(st.jsonText, st.pos, 1)) - 48
            If d >= 0 And d <= 9 Then
                ReadChar st
                hasDigit = True
            Else
                Exit Do
            End If
        Loop
    End If
    If Not hasDigit Then
        Err.Raise ERR_INVALID_JSON, "JsonUtils", "JSON 语法错误: 数字缺少整数部分"
    End If

    ' 可选小数部分 (至少一位数字)
    If PeekChar(st) = "." Then
        ReadChar st
        hasFrac = False
        Do While st.pos <= st.textLen
            d = AscW(Mid$(st.jsonText, st.pos, 1)) - 48
            If d >= 0 And d <= 9 Then
                ReadChar st
                hasFrac = True
            Else
                Exit Do
            End If
        Loop
        If Not hasFrac Then
            Err.Raise ERR_INVALID_JSON, "JsonUtils", "JSON 语法错误: 小数点后缺少数值"
        End If
    End If

    ' 可选指数部分 (至少一位数字)
    If PeekChar(st) = "e" Or PeekChar(st) = "E" Then
        ReadChar st
        If PeekChar(st) = "+" Or PeekChar(st) = "-" Then ReadChar st
        hasExp = False
        Do While st.pos <= st.textLen
            d = AscW(Mid$(st.jsonText, st.pos, 1)) - 48
            If d >= 0 And d <= 9 Then
                ReadChar st
                hasExp = True
            Else
                Exit Do
            End If
        Loop
        If Not hasExp Then
            Err.Raise ERR_INVALID_JSON, "JsonUtils", "JSON 语法错误: 指数部分缺少数字"
        End If
    End If

    numStr = Mid$(st.jsonText, start, st.pos - start)

    ' Val() — 始终以英文句点为小数点，无视系统区域设置
    ' 注: Val() 识别科学计数法 (Val("1.5E+3")=1500)，但为确保跨区域兼容性
    ' 及精确控制指数溢出，手工拆分底数和指数:
    '   Val("1.5E+3") → Val("1.5") * 10^Val("+3") = 1500
    ' (优于 CDbl — CDbl 依赖系统区域设置, 在逗号小数点系统上出错)
    ' 整数若可放入 Long 则保留 Long 类型，否则用 Double
    If InStr(UCase$(numStr), "E") > 0 Then
        ' 手工拆分底数和指数，全程 Val() 保持区域无关
        Dim ePos As Long: ePos = InStr(UCase$(numStr), "E")
        Dim mantissa As Double: mantissa = Val(Left$(numStr, ePos - 1))
        Dim exponent As Double: exponent = Val(Mid$(numStr, ePos + 1))
        ' 防止 10#^exponent 溢出/下溢 (Double 范围 ≈ 10^±308)
        If exponent > 308# Then
            Err.Raise ERR_INVALID_JSON, "JsonUtils", _
                "JSON 数字指数过大: " & numStr
        End If
        If exponent < -324# Then
            ParseNumber = 0#  ' 下溢为零 (与 Double 行为一致)
        Else
            ParseNumber = mantissa * (10# ^ exponent)
        End If
    ElseIf InStr(numStr, ".") > 0 Then
        ParseNumber = Val(numStr)
    Else
        ' 用数值边界取代长度判断，避免 CLng 溢出
        ' (10位十进制数可能远超 Long.MaxValue = 2,147,483,647)
        Dim dVal As Double
        dVal = Val(numStr)
        If dVal >= -2147483648# And dVal <= 2147483647# Then
            ParseNumber = CLng(dVal)
        Else
            ParseNumber = dVal
        End If
    End If
End Function

Private Function ParseLiteral(ByRef st As TJsonState) As Boolean
    Dim remaining As String
    remaining = Mid$(st.jsonText, st.pos, 5)
    If Left$(remaining, 4) = "true" Then
        st.pos = st.pos + 4
        ParseLiteral = True
    ElseIf remaining = "false" Then
        st.pos = st.pos + 5
        ParseLiteral = False
    Else
        Err.Raise ERR_INVALID_JSON, "JsonUtils", _
            "JSON 语法错误: 位置 " & st.pos & " 期望 'true' 或 'false', 实际为 '" & Left$(remaining, 10) & "'"
    End If
End Function

Private Function ParseNull(ByRef st As TJsonState) As Variant
    Dim remaining As String
    remaining = Mid$(st.jsonText, st.pos, 4)
    If remaining = "null" Then
        st.pos = st.pos + 4
        ParseNull = Null
    Else
        Err.Raise ERR_INVALID_JSON, "JsonUtils", _
            "JSON 语法错误: 位置 " & st.pos & " 期望 'null', 实际为 '" & remaining & "'"
    End If
End Function

'=====================================================================
' Public VBA 函数
'=====================================================================

' JsonParse — 解析 JSON 字符串 → Variant
Public Function JsonParse(ByVal json As String) As Variant
    Dim st     As TJsonState
    Dim result As Variant

    If Len(json) = 0 Then
        Err.Raise ERR_INVALID_JSON, "JsonUtils", "JSON 字符串为空。"
    End If

    st.depth = 0  ' 每次解析重置递归深度计数器
    st.jsonText = json
    st.pos = 1
    st.textLen = Len(json)

    ParseValue st, result

    SkipWhitespace st
    If st.pos <= st.textLen Then
        Err.Raise ERR_INVALID_JSON, "JsonUtils", _
            "JSON 语法错误: 位置 " & st.pos & " 多余的字符"
    End If

    ' Inline: function return needs direct Set/Let (VarLetSet fails for arrays)
    If IsObject(result) Then Set JsonParse = result Else JsonParse = result
End Function

' JsonGet — 按路径提取值
Public Function JsonGet(ByVal json As Variant, ByVal path As Variant) As Variant
    ' Normalize path: Range → String
    If IsObject(path) Then
        If TypeOf path Is Range Then path = CStr(path.Value)
    End If
    If IsObject(json) Then
        If TypeOf json Is Range Then json = json.Value
    End If
    If IsArray(json) Then
        On Error Resume Next
        json = json(LBound(json, 1), LBound(json, 2))
        On Error GoTo 0
    End If
    Dim current As Variant
    Dim p       As Long
    Dim c       As String
    Dim brEnd   As Long
    Dim seg     As String
    Dim dotPos  As Long
    Dim brkPos  As Long
    Dim segEnd  As Long

    VarLetSet current, JsonParse(json)

    p = 1
    Do While p <= Len(path)
        c = Mid$(path, p, 1)
        If c = "." Then
            p = p + 1
        ElseIf c = "[" Then
            brEnd = InStr(p, path, "]")
            If brEnd = 0 Then Err.Raise ERR_INVALID_PATH, "JsonUtils", "缺少 ]"
            seg = Mid$(path, p + 1, brEnd - p - 1)
            ' 引号括起的键 → 对象键访问 ["key"] / ['key']
            If (Left$(seg, 1) = """" And Right$(seg, 1) = """") Or _
               (Left$(seg, 1) = "'" And Right$(seg, 1) = "'") Then
                If Not IsObject(current) Then
                    Err.Raise ERR_PATH_NOT_FOUND, "JsonUtils", "期望对象"
                End If
                seg = Mid$(seg, 2, Len(seg) - 2)
                If Not current.Exists(seg) Then
                    Err.Raise ERR_PATH_NOT_FOUND, "JsonUtils", _
                        "键 " & Chr$(39) & seg & Chr$(39) & " 不存在"
                End If
                VarLetSet current, current(seg)
            Else
                ' 数字索引 → 数组访问 [0] / [1]
                If Len(seg) = 0 Or Not IsNumeric(seg) Then
                    Err.Raise ERR_INVALID_PATH, "JsonUtils", _
                        "无效的数组索引 [" & seg & "] — 需要整数"
                End If
                If Not IsArray(current) Then
                    Err.Raise ERR_PATH_NOT_FOUND, "JsonUtils", "期望数组"
                End If
                VarLetSet current, current(CLng(seg))
            End If
            p = brEnd + 1
        Else
            dotPos = InStr(p, path, ".")
            brkPos = InStr(p, path, "[")
            If dotPos = 0 And brkPos = 0 Then
                seg = Mid$(path, p)
                segEnd = Len(path) + 1
            ElseIf dotPos = 0 Then
                seg = Mid$(path, p, brkPos - p)
                segEnd = brkPos
            ElseIf brkPos = 0 Then
                seg = Mid$(path, p, dotPos - p)
                segEnd = dotPos
            Else
                If dotPos < brkPos Then segEnd = dotPos Else segEnd = brkPos
                seg = Mid$(path, p, segEnd - p)
            End If
            If Not IsObject(current) Then
                Err.Raise ERR_PATH_NOT_FOUND, "JsonUtils", "期望对象"
            End If
            If Not current.Exists(seg) Then
                Err.Raise ERR_PATH_NOT_FOUND, "JsonUtils", _
                    "键 " & Chr$(39) & seg & Chr$(39) & " 不存在"
            End If
            VarLetSet current, current(seg)
            p = segEnd
        End If
    Loop

    If IsObject(current) Then Set JsonGet = current Else JsonGet = current
End Function

' JsonToRange — JSON 对象数组 → 工作表
Public Sub JsonToRange(ByVal json As String, ByRef startCell As Range)
    If startCell Is Nothing Then Exit Sub
    If startCell.Areas.Count > 1 Then
        Err.Raise ERR_INVALID_INPUT, "JsonUtils", "不支持多区域 Range，请使用单个连续区域。"
    End If

    Dim root      As Variant
    Dim items     As Variant
    Dim nRows     As Long
    Dim allKeys   As Object
    Dim i         As Long
    Dim j         As Long
    Dim key       As Variant
    Dim nCols     As Long
    Dim keysArr() As String
    Dim outArr()  As Variant
    Dim v         As Variant

    VarLetSet root, JsonParse(json)
    If Not IsArray(root) Then
        Err.Raise ERR_INVALID_JSON, "JsonUtils", "JSON 顶层必须是数组"
    End If

    items = root
    nRows = UBound(items) - LBound(items) + 1
    If nRows = 0 Then Exit Sub

    Set allKeys = DP.Create()
    For i = LBound(items) To UBound(items)
        If IsObject(items(i)) Then
            For Each key In items(i).Keys
                If Not allKeys.Exists(key) Then allKeys.Add key, key
            Next key
        End If
    Next i

    nCols = allKeys.Count
    If nCols = 0 Then Exit Sub

    ReDim keysArr(0 To nCols - 1)
    i = 0
    For Each key In allKeys.Keys
        keysArr(i) = key
        i = i + 1
    Next key

    ReDim outArr(1 To nRows + 1, 1 To nCols)
    For j = 0 To nCols - 1
        outArr(1, j + 1) = keysArr(j)
    Next j

    For i = LBound(items) To UBound(items)
        For j = 0 To nCols - 1
            If IsObject(items(i)) Then
                If items(i).Exists(keysArr(j)) Then
                    VarLetSet v, items(i)(keysArr(j))
                    If IsObject(v) Or IsArray(v) Then
                        outArr(i - LBound(items) + 2, j + 1) = "[嵌套]"
                    ElseIf IsNull(v) Then
                        outArr(i - LBound(items) + 2, j + 1) = Empty
                    Else
                        outArr(i - LBound(items) + 2, j + 1) = v
                    End If
                End If
            End If
        Next j
    Next i

    Dim dest As Range
    Set dest = startCell.Resize(nRows + 1, nCols)
    If dest.HasFormula Or IsNull(dest.HasFormula) Then
        Err.Raise ERR_FORMULA_RANGE, "JsonToRange", _
            "目标区域包含公式，请先使用 ClearContents 清除公式后再写入。"
    End If
    dest.Value = outArr
End Sub

' JsonGetKeys — 获取 JSON 对象所有键
Public Function JsonGetKeys(ByVal json As Variant) As String()
    Dim root      As Variant
    Dim dict      As Object
    Dim result()  As String
    Dim i         As Long
    Dim key       As Variant
    Dim emptyArr() As String

    ' Range→String for cell references (extract first cell safely)
    If IsObject(json) Then
        If TypeOf json Is Range Then
            If json.Count = 1 Then
                json = CStr(json.Value)
            Else
                json = CStr(json.Cells(1, 1).Value)
            End If
        End If
    End If
    ' Guard: if json is still an array/object after Range extraction, bail
    If IsArray(json) Then
        Err.Raise ERR_INVALID_JSON, "JsonUtils", "JSON 输入是数组，无法获取键。请传入 JSON 对象字符串。"
    End If
    VarLetSet root, JsonParse(CStr(json))
    ' Support JSON arrays: return stringified indices as "keys"
    If IsArray(root) Then
        Dim ai As Long, alo As Long, ahi As Long
        alo = LBound(root): ahi = UBound(root)
        ReDim result(alo To ahi) As String
        For ai = alo To ahi
            result(ai) = CStr(ai)
        Next ai
        JsonGetKeys = result
        Exit Function
    End If
    If Not IsObject(root) Then
        Err.Raise ERR_INVALID_JSON, "JsonUtils", _
            "JSON 顶层不是对象也不是数组 (类型: " & TypeName(root) & ")。"
    End If

    Set dict = root
    If dict.Count = 0 Then
        Erase emptyArr
        JsonGetKeys = emptyArr
        Exit Function
    End If

    ReDim result(0 To dict.Count - 1)
    i = 0
    For Each key In dict.Keys
        result(i) = key
        i = i + 1
    Next key
    JsonGetKeys = result
End Function

' JsonIsValid — JSON 合法性校验 (不抛错)
Public Function JsonIsValid(ByVal json As Variant) As Boolean
    If IsObject(json) Then
        If TypeOf json Is Range Then json = json.Value
    End If
    If Len(CStr(json)) = 0 Then Exit Function
    Err.Clear: On Error Resume Next
    JsonParse CStr(json)
    JsonIsValid = (Err.Number = 0)
    On Error GoTo 0
End Function

'=====================================================================
' JsonStringify — 将 VBA Variant 序列化为 JSON 字符串
'
' 支持: Dictionary → 对象, Array → 数组, Null/Boolean/Number/String/标量
' 输出: 紧凑 JSON (无缩进), 键序由 Dictionary.Keys 决定 (非确定性)
' 字符串转义: " \ / 控制字符 (b f n r t uXXXX) → 标准 JSON 转义
' 数值: Str$ 后 Replace(",",".") 确保小数点始终为 "." — 与系统区域设置无关
'
' 限制:
'   - 不检测循环引用 (递归深度上限约 100 层)
'   - 日期 → ISO 8601 字符串 ("2026-06-07T...")
'   - 空值 / 未初始化 → null
'=====================================================================
Public Function JsonStringify(ByVal value As Variant) As String
    JsonStringify = StringifyValue(value, 0)
End Function

' StringifyValue — 递归序列化核心 (Private)
' depth: 递归深度计数器, 超过 100 层返回 "[[Circular]]" (#28)
Private Function StringifyValue(ByVal v As Variant, Optional ByVal depth As Long) As String
    Const MAX_DEPTH As Long = 100
    If depth > MAX_DEPTH Then
        StringifyValue = """[[Circular]]"""
        Exit Function
    End If
    Dim key As Variant, i As Long, n As Long
    Dim parts() As String, pIdx As Long, first As Boolean
    Dim dict As Object

    If IsObject(v) Then
        If TypeName(v) <> "Dictionary" Then
            StringifyValue = """" & TypeName(v) & """": Exit Function
        End If
        Set dict = v
        If dict.Count = 0 Then StringifyValue = "{}": Exit Function
        ReDim parts(0 To dict.Count * 4)
        pIdx = 0: parts(0) = "{": pIdx = 1: first = True
        For Each key In dict.Keys
            If Not first Then parts(pIdx) = ",": pIdx = pIdx + 1
            first = False
            parts(pIdx) = """" & EscapeJSONString(CStr(key)) & """:": pIdx = pIdx + 1
            parts(pIdx) = StringifyValue(dict(key), depth + 1): pIdx = pIdx + 1
        Next
        parts(pIdx) = "}": pIdx = pIdx + 1
        ReDim Preserve parts(0 To pIdx - 1)
        StringifyValue = Join(parts, "")
        Exit Function
    End If

    If IsArray(v) Then
        ' Detect dimensions: try UBound on dim 2 (§9.3: Err.Clear 前置清理上游残留)
        Dim is2D As Boolean
        Err.Clear: On Error Resume Next
        Dim testUB As Long: testUB = UBound(v, 2)
        is2D = (Err.Number = 0)
        On Error GoTo 0

        If is2D Then
            ' 2D array → serialize as array of rows
            Dim nRows As Long: nRows = UBound(v, 1) - LBound(v, 1) + 1
            Dim nCols As Long: nCols = UBound(v, 2) - LBound(v, 2) + 1
            If nRows <= 0 Or nCols <= 0 Then StringifyValue = "[]": Exit Function
            ReDim parts(0 To nRows * 4)
            pIdx = 0: parts(0) = "[": pIdx = 1: first = True
            Dim r As Long, c As Long
            For r = LBound(v, 1) To UBound(v, 1)
                If Not first Then parts(pIdx) = ",": pIdx = pIdx + 1
                first = False
                ' Build row as sub-array string
                Dim rowParts() As String, rpIdx As Long, rFirst As Boolean
                ReDim rowParts(0 To nCols * 2)
                rpIdx = 0: rowParts(0) = "[": rpIdx = 1: rFirst = True
                For c = LBound(v, 2) To UBound(v, 2)
                    If Not rFirst Then rowParts(rpIdx) = ",": rpIdx = rpIdx + 1
                    rFirst = False
                    rowParts(rpIdx) = StringifyValue(v(r, c), depth + 1): rpIdx = rpIdx + 1
                Next
                rowParts(rpIdx) = "]": rpIdx = rpIdx + 1
                ReDim Preserve rowParts(0 To rpIdx - 1)
                parts(pIdx) = Join(rowParts, ""): pIdx = pIdx + 1
            Next
            parts(pIdx) = "]": pIdx = pIdx + 1
        Else
            ' 1D array
            Dim lb As Long: lb = LBound(v)
            Dim ub As Long: ub = UBound(v)
            n = ub - lb + 1
            If n <= 0 Then StringifyValue = "[]": Exit Function
            ReDim parts(0 To n * 4)
            pIdx = 0: parts(0) = "[": pIdx = 1: first = True
            For i = lb To ub
                If Not first Then parts(pIdx) = ",": pIdx = pIdx + 1
                first = False
                parts(pIdx) = StringifyValue(v(i), depth + 1): pIdx = pIdx + 1
            Next
            parts(pIdx) = "]": pIdx = pIdx + 1
        End If

        ReDim Preserve parts(0 To pIdx - 1)
        StringifyValue = Join(parts, "")
        Exit Function
    End If

    If IsNull(v) Or IsEmpty(v) Then StringifyValue = "null": Exit Function
    If VarType(v) = vbBoolean Then
        If v Then StringifyValue = "true" Else StringifyValue = "false"
        Exit Function
    End If
    If VarType(v) = vbDate Then
        StringifyValue = """" & Format$(v, "yyyy-mm-dd\Thh:nn:ss") & """": Exit Function
    End If
    If VarType(v) <> vbString And IsNumeric(v) Then
        Dim d As Double: d = CDbl(v)
        ' 整数在 Long 范围内输出无小数点
        If Abs(d) <= 2147483647# Then
            If d = CLng(d) Then
                StringifyValue = CStr(CLng(d))
            Else
                StringifyValue = Replace$(LTrim$(Str$(d)), ",", ".")
            End If
        Else
            StringifyValue = Replace$(LTrim$(Str$(d)), ",", ".")
        End If
        Exit Function
    End If

    ' Fallback: 字符串 — 需转义
    StringifyValue = """" & EscapeJSONString(CStr(v)) & """"
End Function

' EscapeJSONString — JSON 字符串转义 (Private)
Private Function EscapeJSONString(ByVal s As String) As String
    Dim i As Long, n As Long, ch As String, code As Long
    Dim parts() As String, pIdx As Long

    n = Len(s)
    If n = 0 Then EscapeJSONString = "": Exit Function

    ReDim parts(0 To n - 1)
    pIdx = 0
    For i = 1 To n
        ch = Mid$(s, i, 1): code = AscW(ch)
        Select Case code
            Case 34: parts(pIdx) = "\""": pIdx = pIdx + 1   ' "
            Case 92: parts(pIdx) = "\\": pIdx = pIdx + 1    ' \
            Case 47: parts(pIdx) = "\/": pIdx = pIdx + 1    ' /
            Case 8:  parts(pIdx) = "\b": pIdx = pIdx + 1    ' BS
            Case 12: parts(pIdx) = "\f": pIdx = pIdx + 1    ' FF
            Case 10: parts(pIdx) = "\n": pIdx = pIdx + 1    ' LF
            Case 13: parts(pIdx) = "\r": pIdx = pIdx + 1    ' CR
            Case 9:  parts(pIdx) = "\t": pIdx = pIdx + 1    ' TAB
            Case Else
                If code < 32 Then
                    parts(pIdx) = "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    parts(pIdx) = ch
                End If
                pIdx = pIdx + 1
        End Select
    Next
    If pIdx > 0 Then ReDim Preserve parts(0 To pIdx - 1)
    EscapeJSONString = Join(parts, "")
End Function

'=====================================================================
' 工作表函数 (UDF_JSON_*)
'=====================================================================

Public Function UDF_JSON_GET(ByVal json As Variant, ByVal path As Variant) As Variant
    On Error GoTo EH
    Dim result As Variant
    VarLetSet result, JsonGet(json, path)
    If IsObject(result) Then UDF_JSON_GET = "[对象]": Exit Function
    If IsArray(result) Then UDF_JSON_GET = "[数组]": Exit Function
    If IsNull(result) Then UDF_JSON_GET = "null": Exit Function
    UDF_JSON_GET = result
    Exit Function
EH:
    UDF_JSON_GET = CVErr(xlErrValue)
End Function

Public Function UDF_JSON_KEYS(ByVal json As Variant) As Variant
    On Error GoTo EH
    Dim keys()   As String
    Dim i        As Long
    Dim n        As Long
    Dim outArr() As Variant

    keys = JsonGetKeys(json)
    Dim isEmptyArr As Boolean
    On Error Resume Next
    Dim lbTest As Long: lbTest = LBound(keys)
    If Err.Number <> 0 Then
        Err.Clear: isEmptyArr = True
    Else
        Err.Clear: isEmptyArr = (UBound(keys) < LBound(keys))
    End If
    On Error GoTo 0
    If isEmptyArr Then
        UDF_JSON_KEYS = CVErr(xlErrNA)
        Exit Function
    End If
    n = UBound(keys) - LBound(keys) + 1
    If n <= 0 Then
        UDF_JSON_KEYS = CVErr(xlErrNA)
        Exit Function
    End If

    ReDim outArr(1 To n, 1 To 1)
    For i = 0 To n - 1
        outArr(i + 1, 1) = keys(i)
    Next i
    UDF_JSON_KEYS = outArr
    Exit Function
EH:
    UDF_JSON_KEYS = CVErr(xlErrValue)
End Function

Public Function UDF_JSON_IS_VALID(ByVal json As Variant) As Variant
    On Error GoTo EH
    UDF_JSON_IS_VALID = JsonIsValid(json)
    Exit Function
EH:
    UDF_JSON_IS_VALID = CVErr(xlErrValue)
End Function

Public Function UDF_JSON_STRINGIFY(ByVal value As Variant) As Variant
    On Error GoTo EH
    UDF_JSON_STRINGIFY = JsonStringify(value)
    Exit Function
EH:
    UDF_JSON_STRINGIFY = CVErr(xlErrValue)
End Function