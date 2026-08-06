Option Explicit

'==============================================================================
' Module:       RangeUtils
' Purpose:      Range export: HTML/JSON/MD/CSV, area operations, naming
' Layer:        Excel
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       49 functions/subs
'==============================================================================


'=====================================================================
' RangeUtils.bas — 工作表区域工具集 (通用标准库)

Private Const ERR_MULTI_AREA As Long = vbObjectError + 1005
Private Const ERR_MULTI_SELECT As Long = vbObjectError + 1003
Private Const ERR_STREAM_FAIL As Long = vbObjectError + 1010
Private Const ERR_FORMULA_RANGE As Long = vbObjectError + 1014
'
' 注意: 本模块函数不支持多区域选择 (Multi-Area Range)。
'       除 CountVisible / RangeDiff / FindAll 显式处理多区域外，
'       传入包含多个 Area 的 Range 将导致运行时错误 1004。
'       FindAll 仅搜索第一个 Area，上限 8,000 个匹配单元格。
'
' 错误码:
'   vbObjectError + 1005 — 多区域/尺寸不匹配
'
' 区域定位:
'   LastRow / LastCol — 工作表最后数据行/列号
'   ColLetter / ColNumber — 列号 ↔ 字母
'   UsedRangeEx      — 实际数据区域
'   GetCellAddress   — 构建单元格地址字符串
'   IsRangeEmpty     — 区域是否为空
'   CountVisible     — 可见行数
'   RangeExists      — 引用是否为 #REF!
'   NamedRangeExists — 名称是否存在
'
' 命名区域:
'   NamedRangeAdd    — 创建/更新命名区域
'   NamedRangeDelete — 删除命名区域
'
' 区域操作:
'   TransposeRange   — 转置粘贴
'   RemoveDuplicatesRange — 按列去重
'   SortRange        — 按列排序
'   FilterRangeToArray — 条件筛选返回数组
'   RangeToArray     — 区域→二维数组 (无筛选)
'   FindAll          — 查找所有匹配
'   MergeRanges      — 合并多区域 (并集)
'   IntersectRanges  — 取区域交集
'   RangeDiff        — 比较区域差异
'
' 区域维护:
'   ClearRangeContents — 清除内容
'   ClearRangeFormats  — 清除格式
'   AutoFitRange       — 自动调整列宽/行高
'
' 导出/转换:
'   RangeToHTML      — → HTML 表格
'   RangeToJSON      — → JSON 数组
'   RangeToMarkdown  — → Markdown 表格
'   ExportRangeToCSV — → CSV 文件
'   CopyRangeToSheet — 数组 → 工作表
'   SafeText         — 安全转换单元格值为字符串
'
' 工作表函数 (UDF_RANGE_*):
'   UDF_RANGE_LASTROW/FIRSTROW/LASTCOL/FIRSTCOL — 数据边界
'   UDF_RANGE_COL_LETTER/COL_NUM — 列号↔字母
'   UDF_RANGE_EXISTS/NAMEDEXISTS/ISEMPTY/COUNTVISIBLE — 检测
'   UDF_RANGE_CELLADDRESS/SAFETEXT — 转换
'   UDF_RANGE_TOHTML/TOJSON/TOMD — 导出
'   UDF_RANGE_FILTER — 条件筛选
'=====================================================================

'=============================================================================
' LastRow — 获取工作表最后有数据的行号
'=============================================================================
Public Function LastRow(Optional ByVal ws As Worksheet = Nothing) As Long
    Dim wsLocal As Worksheet
    Set wsLocal = ws
    If wsLocal Is Nothing Then
        On Error Resume Next
        Set wsLocal = Application.ActiveSheet
        On Error GoTo 0
        If wsLocal Is Nothing Then Exit Function
    End If

    On Error Resume Next
    LastRow = wsLocal.Cells.Find( _
        What:="*", _
        After:=wsLocal.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Row
    If Err.Number <> 0 Then LastRow = 0
    On Error GoTo 0
End Function

'=============================================================================
' LastCol — 获取工作表最后有数据的列号
'=============================================================================
Public Function LastCol(Optional ByVal ws As Worksheet = Nothing) As Long
    Dim wsLocal As Worksheet
    Set wsLocal = ws
    If wsLocal Is Nothing Then
        On Error Resume Next
        Set wsLocal = Application.ActiveSheet
        On Error GoTo 0
        If wsLocal Is Nothing Then Exit Function
    End If

    On Error Resume Next
    LastCol = wsLocal.Cells.Find( _
        What:="*", _
        After:=wsLocal.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Column
    If Err.Number <> 0 Then LastCol = 0
    On Error GoTo 0
End Function

'=============================================================================
' FirstRow — 获取工作表最先有数据的行号
'=============================================================================
Public Function FirstRow(Optional ByVal ws As Worksheet = Nothing) As Long
    Dim wsLocal As Worksheet
    Set wsLocal = ws
    If wsLocal Is Nothing Then
        On Error Resume Next
        Set wsLocal = Application.ActiveSheet
        On Error GoTo 0
        If wsLocal Is Nothing Then Exit Function
    End If

    On Error Resume Next
    FirstRow = wsLocal.Cells.Find( _
        What:="*", _
        After:=wsLocal.Cells(wsLocal.Rows.Count, wsLocal.Columns.Count), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False).Row
    If Err.Number <> 0 Then FirstRow = 0
    On Error GoTo 0
End Function

'=============================================================================
' FirstCol — 获取工作表最先有数据的列号
'=============================================================================
Public Function FirstCol(Optional ByVal ws As Worksheet = Nothing) As Long
    Dim wsLocal As Worksheet
    Set wsLocal = ws
    If wsLocal Is Nothing Then
        On Error Resume Next
        Set wsLocal = Application.ActiveSheet
        On Error GoTo 0
        If wsLocal Is Nothing Then Exit Function
    End If

    On Error Resume Next
    FirstCol = wsLocal.Cells.Find( _
        What:="*", _
        After:=wsLocal.Cells(wsLocal.Rows.Count, wsLocal.Columns.Count), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, _
        MatchCase:=False).Column
    If Err.Number <> 0 Then FirstCol = 0
    On Error GoTo 0
End Function

'=============================================================================
' ColLetter — 列号 → 字母: 1→"A", 27→"AA", 703→"AAA"
'=============================================================================
Public Function ColLetter(ByVal colNum As Long) As String
    If colNum < 1 Then
        ColLetter = ""
        Exit Function
    End If

    Dim result As String
    Dim remaining As Long
    Dim idx As Long
    remaining = colNum

    Do While remaining > 0
        idx = ((remaining - 1) Mod 26) + 1
        result = Chr$(64 + idx) & result
        remaining = (remaining - 1) \ 26
    Loop
    ColLetter = result
End Function

'=============================================================================
' ColNumber — 字母 → 列号: "A"→1, "AA"→27
'=============================================================================
Public Function ColNumber(ByVal colRef As Variant) As Long
    Dim colStr As String
    Dim i As Long, ch As Long

    If IsError(colRef) Then
        ColNumber = -1
        Exit Function
    End If

    ' 仅将非字符串数值视为直接的列号
    If VarType(colRef) <> vbString Then
        If IsNumeric(colRef) Then
            ColNumber = CLng(colRef)
            Exit Function
        End If
    End If

    colStr = UCase$(Trim(CStr(colRef)))
    If Len(colStr) = 0 Then
        ColNumber = -1
        Exit Function
    End If

    For i = 1 To Len(colStr)
        ch = AscW(Mid$(colStr, i, 1))
        If ch < 65 Or ch > 90 Then
            ColNumber = -1
            Exit Function
        End If
        ColNumber = ColNumber * 26 + (ch - 64)
    Next i
End Function

'=============================================================================
' UsedRangeEx — 获取实际数据区域
'
' 比内置 UsedRange 更健壮 (内置 UsedRange 可能包含格式化但无数据的空单元格)
'=============================================================================
Public Function UsedRangeEx(Optional ByVal ws As Worksheet = Nothing) As Range
    Dim wsLocal As Worksheet
    Set wsLocal = ws
    If wsLocal Is Nothing Then
        On Error Resume Next
        Set wsLocal = Application.ActiveSheet
        On Error GoTo 0
        If wsLocal Is Nothing Then Exit Function
    End If

    Dim fr As Long, fc As Long, lr As Long, lc As Long
    fr = FirstRow(wsLocal)
    fc = FirstCol(wsLocal)
    lr = LastRow(wsLocal)
    lc = LastCol(wsLocal)

    If fr = 0 Or fc = 0 Or lr = 0 Or lc = 0 Then
        Set UsedRangeEx = Nothing
    Else
        Set UsedRangeEx = wsLocal.Range(wsLocal.Cells(fr, fc), wsLocal.Cells(lr, lc))
    End If
End Function

'=============================================================================
' RangeToHTML — 区域转 HTML 表格
'
' 参数:
'   rng           - 源区域
'   includeHeader - 是否将第一行作为 <thead> (默认 True)
'   tableClass    - 可选 CSS class 属性
'=============================================================================
Public Function RangeToHTML( _
    ByRef rng As Range, _
    Optional ByVal includeHeader As Boolean = True, _
    Optional ByVal tableClass As String = "") As String

    Dim data As Variant
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim startRow As Long
    Dim clsAttr As String
    Dim parts() As String
    Dim pIdx As Long

    If rng Is Nothing Then
        RangeToHTML = ""
        Exit Function
    End If

    GetRangeData rng, data, nRows, nCols

    If Len(tableClass) > 0 Then
        clsAttr = " class=""" & EscapeHTML(tableClass) & """"
    End If

    pIdx = 0
    ReDim parts(0 To nRows * (nCols * 3 + 4) + 10)

    parts(pIdx) = "<table" & clsAttr & ">": pIdx = pIdx + 1

    If includeHeader And nRows >= 1 Then
        parts(pIdx) = "<thead><tr>": pIdx = pIdx + 1
        For j = 1 To nCols
            parts(pIdx) = "<th>" & EscapeHTML(SafeText(data(1, j))) & "</th>"
            pIdx = pIdx + 1
        Next j
        parts(pIdx) = "</tr></thead><tbody>": pIdx = pIdx + 1
        startRow = 2
    Else
        parts(pIdx) = "<tbody>": pIdx = pIdx + 1
        startRow = 1
    End If

    For i = startRow To nRows
        EnsureCapacity parts, pIdx, nCols * 3 + 2
        parts(pIdx) = "<tr>": pIdx = pIdx + 1
        For j = 1 To nCols
            parts(pIdx) = "<td>" & EscapeHTML(SafeText(data(i, j))) & "</td>"
            pIdx = pIdx + 1
        Next j
        parts(pIdx) = "</tr>": pIdx = pIdx + 1
    Next i

    parts(pIdx) = "</tbody></table>"
    If pIdx > 0 Then
        ReDim Preserve parts(0 To pIdx)
        RangeToHTML = Join(parts, "")
    End If
End Function

Public Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then
        SafeText = "[Error]"
    ElseIf IsNull(v) Then
        SafeText = ""
    ElseIf IsEmpty(v) Then
        SafeText = ""
    ElseIf IsObject(v) Then
        If TypeOf v Is Range Then
            SafeText = CStr(v.Cells(1, 1).Value2)
        Else
            SafeText = "[Object]"
        End If
    Else
        SafeText = CStr(v)
    End If
End Function

Private Function EscapeHTML(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, """", "&quot;")
    EscapeHTML = text
End Function

Private Sub GetRangeData(ByRef rng As Range, ByRef data As Variant, ByRef nRows As Long, ByRef nCols As Long)
    If rng Is Nothing Then
        nRows = 0: nCols = 0: Exit Sub
    End If
    If rng.Count = 0 Then
        ' 防御: 空区域 (COM 边界情况 / 损坏引用) — 返回 0×0，由调用方检查 nRows
        nRows = 0: nCols = 0: Exit Sub
    End If
    If rng.Areas.Count > 1 Then
        Err.Raise ERR_MULTI_SELECT, "GetRangeData", "不支持多选区，请使用连续区域。"
    End If
    If rng.Count = 1 Then
        ReDim data(1 To 1, 1 To 1)
        data(1, 1) = rng.Value
        nRows = 1: nCols = 1
    Else
        data = rng.Value
        nRows = UBound(data, 1)
        nCols = UBound(data, 2)
    End If
End Sub

Private Sub EnsureCapacity(ByRef parts() As String, ByVal pIdx As Long, ByVal needed As Long)
    If pIdx + needed > UBound(parts) Then
        ReDim Preserve parts(0 To UBound(parts) * 2)
    End If
End Sub

'=============================================================================
' FindAll — 查找所有匹配单元格
'
' 返回: Range 集合 (可迭代), 未找到返回 Nothing
'
' Excel Application.Union 上限 ~8,000 个 Area；超出时 outTruncated = True。
'=============================================================================
Public Function FindAll( _
    ByRef rng As Range, _
    ByVal value As Variant, _
    Optional ByVal lookIn As XlFindLookIn = xlValues, _
    Optional ByVal lookAt As XlLookAt = xlPart, _
    Optional ByVal matchCase As Boolean = False, _
    Optional ByRef outTruncated As Boolean) As Range

    Dim firstAddress As String
    Dim cell As Range
    Dim result As Range

    outTruncated = False

    ' 注意：多区域 Range 仅搜索第一个 Area。如需支持多区域，
    ' 请对 rng.Areas 循环调用 FindAll。
    If rng Is Nothing Then
        Set FindAll = Nothing
        Exit Function
    End If

    Set cell = rng.Find( _
        What:=value, _
        LookIn:=lookIn, _
        LookAt:=lookAt, _
        MatchCase:=matchCase)

    If cell Is Nothing Then
        Set FindAll = Nothing
        Exit Function
    End If

    firstAddress = cell.Address
    Set result = cell
    Const MAX_UNION_AREAS As Long = 8000

    Do
        Set cell = rng.FindNext(cell)
        If cell Is Nothing Then Exit Do
        If cell.Address = firstAddress Then Exit Do
        If result.Areas.Count < MAX_UNION_AREAS Then
            Set result = Application.Union(result, cell)
        Else
            outTruncated = True
            Exit Do
        End If
    Loop

    Set FindAll = result
End Function

'=============================================================================
' MergeRanges — 合并多个不连续区域
'=============================================================================
Public Function MergeRanges(ParamArray ranges() As Variant) As Range
    Dim result As Range
    Dim i As Long
    Dim first As Boolean
    Dim r As Range

    first = True
    For i = LBound(ranges) To UBound(ranges)
        If TypeName(ranges(i)) = "Range" Then
            Set r = ranges(i)
            If Not r Is Nothing Then
                If first Then
                    Set result = r
                    first = False
                ElseIf result.Parent Is r.Parent Then
                    Set result = Application.Union(result, r)
                End If
            End If
        End If
    Next i

    If first Then
        Set MergeRanges = Nothing
    Else
        Set MergeRanges = result
    End If
End Function

'=============================================================================
' IntersectRanges — 取区域交集
'
' 返回: 所有区域的公共部分, 无交集返回 Nothing
'=============================================================================
Public Function IntersectRanges(ParamArray ranges() As Variant) As Range
    Dim i As Long
    Dim result As Range
    Dim r As Range

    For i = LBound(ranges) To UBound(ranges)
        If TypeName(ranges(i)) = "Range" Then
            Set r = ranges(i)
            If Not r Is Nothing Then
                If result Is Nothing Then
                    Set result = r
                ElseIf result.Parent Is r.Parent Then
                    Set result = Application.Intersect(result, r)
                    If result Is Nothing Then Exit Function
                Else
                    Set IntersectRanges = Nothing
                    Exit Function
                End If
            End If
        End If
    Next i
    Set IntersectRanges = result
End Function


'=============================================================================
' ExportRangeToCSV — 导出区域为 CSV 文件
'
' 参数:
'   rng           - 源区域
'   filePath      - 输出路径
'   delimiter     - 分隔符 (默认 ",")
'   bom           - 是否添加 BOM (默认 True); UTF-8/UTF-16LE 下 ADODB.Stream 保存时自动写 BOM,
'                   bom=False 时通过二进制复制跳过自动 BOM (UTF-8: 3 字节, UTF-16LE: 2 字节)
'   enc           - 编码: "UTF-8", "ANSI", "UTF-16LE"
'   includeHeader - 是否包含第一行 (默认 True)
'=============================================================================
Public Sub ExportRangeToCSV( _
    ByRef rng As Range, _
    ByVal filePath As String, _
    Optional ByVal delimiter As String = ",", _
    Optional ByVal bom As Boolean = True, _
    Optional ByVal enc As String = "UTF-8", _
    Optional ByVal includeHeader As Boolean = True)

    Dim data As Variant
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim startRow As Long
    Dim rowParts() As String
    Dim rpIdx As Long

    If rng Is Nothing Then Exit Sub
    If Len(filePath) = 0 Then Exit Sub

    GetRangeData rng, data, nRows, nCols

    Dim adoStream As Object
    Dim rawStream As Object
    Dim bomLen As Long
    ' 此处需要使用标签式错误处理：ADODB.Stream 即使失败也必须关闭。
    ' 其他函数使用 Resume Next，因为它们没有外部资源需要清理。
    On Error GoTo StreamError
    Set adoStream = CreateObject("ADODB.Stream")
    With adoStream
        .Type = 2 ' 文本类型

        Select Case UCase$(enc)
            Case "UTF-16LE", "UNICODE"
                .Charset = "UTF-16LE"
            Case "UTF-8"
                .Charset = "UTF-8"
            Case "ANSI", "ASCII", "WINDOWS-1252"
                .Charset = "Windows-1252"
            Case Else
                .Charset = enc
        End Select
        .Open

        ' 注意: ADODB.Stream 保存 UTF-8/UTF-16LE 时自动写入 BOM — 不得手工再写 (避免双 BOM)

        If includeHeader Then startRow = 1 Else startRow = 2

        ReDim rowParts(0 To nCols * 2)

        For i = startRow To nRows
            rpIdx = 0
            For j = 1 To nCols
                If j > 1 Then
                    rowParts(rpIdx) = delimiter: rpIdx = rpIdx + 1
                End If
                rowParts(rpIdx) = CSVEncode(SafeText(data(i, j)), delimiter)
                rpIdx = rpIdx + 1
            Next j
            ReDim Preserve rowParts(0 To rpIdx - 1)
            .WriteText Join(rowParts, ""), 1 ' 写入行
        Next i

        ' bom=False: 跳过 ADODB.Stream 自动写入的 BOM (二进制复制)
        bomLen = 0
        If Not bom Then
            Select Case UCase$(enc)
                Case "UTF-8": bomLen = 3
                Case "UTF-16LE", "UNICODE": bomLen = 2
            End Select
        End If
        If bomLen > 0 Then
            Set rawStream = CreateObject("ADODB.Stream")
            rawStream.Type = 1 ' binary
            rawStream.Open
            .Position = bomLen
            .CopyTo rawStream
            rawStream.SaveToFile filePath, 2 ' 保存覆盖
            rawStream.Close
            Set rawStream = Nothing
        Else
            .SaveToFile filePath, 2 ' 保存覆盖
        End If
        .Close
    End With
    Set adoStream = Nothing
    Exit Sub

StreamError:
    ' 清理 ADODB 资源
    If Not rawStream Is Nothing Then
        On Error Resume Next
        rawStream.Close
        Set rawStream = Nothing
        On Error GoTo 0
    End If
    If Not adoStream Is Nothing Then
        On Error Resume Next
        adoStream.Close
        Set adoStream = Nothing
        On Error GoTo 0
    End If

    ' 纯 VBA 回退: Open For Output（ANSI 编码，无 BOM）
    Dim fNum As Integer
    On Error GoTo FallbackFail
    fNum = FreeFile
    Open filePath For Output As #fNum
    ' 回退路径仅产生 ANSI 输出 — 写入 BOM 会因 Chr$ 对代码页的依赖而损坏
    ' (Chr$(&HEF) 在 CP936 等 DBCS 代码页下映射到非 BOM 字节)
    ' If bom Then Print #fNum, Chr$(&HEF) & Chr$(&HBB) & Chr$(&HBF);
    If includeHeader Then startRow = 1 Else startRow = 2
    For i = startRow To nRows
        rpIdx = 0
        ReDim rowParts(0 To nCols * 2)
        For j = 1 To nCols
            If j > 1 Then rowParts(rpIdx) = delimiter: rpIdx = rpIdx + 1
            rowParts(rpIdx) = CSVEncode(SafeText(data(i, j)), delimiter): rpIdx = rpIdx + 1
        Next j
        ReDim Preserve rowParts(0 To rpIdx - 1)
        Print #fNum, Join(rowParts, "")
    Next i
    Close #fNum
    Exit Sub

FallbackFail:
    If fNum > 0 Then On Error Resume Next: Close #fNum: On Error GoTo 0
    Err.Raise ERR_STREAM_FAIL, "ExportRangeToCSV", "ADODB.Stream 不可用且纯 VBA 回退也失败。"
End Sub

Private Function CSVEncode(ByVal text As String, ByVal delimiter As String) As String
    If InStr(text, delimiter) > 0 Or InStr(text, """") > 0 Or _
       InStr(text, vbLf) > 0 Or InStr(text, vbCr) > 0 Then
        CSVEncode = """" & Replace(text, """", """""") & """"
    Else
        CSVEncode = text
    End If
End Function
'=============================================================================
' RangeToJSON — 区域转 JSON 数组
'
' includeHeader=True 时: [{"col1":"val",...}, ...]
' includeHeader=False 时: [["val1","val2"], ...]
'=============================================================================
Public Function RangeToJSON( _
    ByRef rng As Range, _
    Optional ByVal includeHeader As Boolean = True) As String

    Dim data As Variant
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim headers() As String
    Dim parts() As String
    Dim pIdx As Long
    Dim startRow As Long

    If rng Is Nothing Then
        RangeToJSON = "[]"
        Exit Function
    End If

    If rng.Count = 1 Then
        If includeHeader Then
            RangeToJSON = "[{""Value"":" & JSONValue(rng.Value) & "}]"
        Else
            RangeToJSON = "[[" & JSONValue(rng.Value) & "]]"
        End If
        Exit Function
    End If

    If rng.Areas.Count > 1 Then
        Err.Raise ERR_MULTI_SELECT, "RangeToJSON", "不支持多选区，请使用连续区域。"
    End If
    data = rng.Value
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)

    ' 提取标题
    If includeHeader Then
        ReDim headers(1 To nCols)
        For j = 1 To nCols
            headers(j) = SafeText(data(1, j))
        Next j
    End If

    pIdx = 0
    ' 预估: 每个单元格一行 JSON 片段约 nCols*6 个片段
    ReDim parts(0 To (nRows + 1) * nCols * 8 + nRows + 10)

    parts(pIdx) = "[": pIdx = pIdx + 1
    If includeHeader Then startRow = 2 Else startRow = 1

    For i = startRow To nRows
        EnsureCapacity parts, pIdx, nCols * 3 + 5
        If i > startRow Then
            parts(pIdx) = ",": pIdx = pIdx + 1
        End If

        If includeHeader Then
            ' 对象格式: {"key1":"val1","key2":"val2"}
            parts(pIdx) = "{": pIdx = pIdx + 1
            For j = 1 To nCols
                If j > 1 Then
                    parts(pIdx) = ",": pIdx = pIdx + 1
                End If
                parts(pIdx) = """" & JSONEscape(headers(j)) & """:"
                parts(pIdx + 1) = JSONValue(data(i, j))
                pIdx = pIdx + 2
            Next j
            parts(pIdx) = "}": pIdx = pIdx + 1
        Else
            ' 数组格式: ["val1","val2"]
            parts(pIdx) = "[": pIdx = pIdx + 1
            For j = 1 To nCols
                If j > 1 Then
                    parts(pIdx) = ",": pIdx = pIdx + 1
                End If
                parts(pIdx) = JSONValue(data(i, j))
                pIdx = pIdx + 1
            Next j
            parts(pIdx) = "]": pIdx = pIdx + 1
        End If
    Next i
    parts(pIdx) = "]"

    ReDim Preserve parts(0 To pIdx)
    RangeToJSON = Join(parts, "")
End Function

Private Function JSONEscape(ByVal text As String) As String
    text = Replace(text, "\", "\\")
    text = Replace(text, """", "\""")
    text = Replace(text, vbCr, "\r")
    text = Replace(text, vbLf, "\n")
    text = Replace(text, vbTab, "\t")
    text = Replace(text, vbBack, "\b")
    text = Replace(text, vbFormFeed, "\f")
    text = Replace(text, vbNullChar, "\u0000")
    text = Replace(text, ChrW$(11), "\u000b")
    JSONEscape = text
End Function

Private Function JSONValue(ByVal v As Variant) As String
    ' 优先从 Variant 子类型检测类型，回退到字符串解析
    ' 避免了 CStr 的区域设置问题，并正确处理 VBA 的 Date/Double/Boolean 类型
    Dim vType As Long, text As String, d As Double, dt As Date

    If IsError(v) Then
        JSONValue = """[Error]"""
        Exit Function
    End If
    If IsNull(v) Then
        JSONValue = "null"
        Exit Function
    End If
    If IsEmpty(v) Then
        JSONValue = """"""
        Exit Function
    End If

    vType = VarType(v)

    If vType = vbBoolean Then
        If v Then JSONValue = "true" Else JSONValue = "false"
        Exit Function
    End If

    If vType = vbDate Then
        JSONValue = """" & Format$(v, "yyyy-mm-dd""T""hh:nn:ss") & """"
        Exit Function
    End If

    ' 数值子类型: vbInteger(2)、vbLong(3)、vbSingle(4)、vbDouble(5)、
    ' vbCurrency(6)、vbDecimal(14)、vbByte(17)、vbLongLong(20)
    ' VarType 枚举值不连续 — 必须使用 Select Case，不能做范围检查
    Select Case vType
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, vbLongLong
            If Int(v) = v Then
                ' 仅当值在 Long 范围内使用 CLng，否则使用 Str$ 格式化
                If CDbl(v) >= -2147483648# And CDbl(v) <= 2147483647# Then
                    JSONValue = CStr(CLng(v))
                Else
                    JSONValue = Replace$(LTrim$(Str$(CDbl(v))), ",", ".")
                End If
            Else
                ' Str$ 后 Replace(",",".") 确保 JSON 合规 — 与系统区域设置无关
                JSONValue = Replace$(LTrim$(Str$(CDbl(v))), ",", ".")
            End If
            Exit Function
    End Select

    ' 回退：基于字符串的类型检测，用于看似数字/日期的文本
    text = CStr(v)
    If Len(text) = 0 Then
        JSONValue = """"""
        Exit Function
    End If

    Select Case LCase$(text)
        Case "true":  JSONValue = "true":  Exit Function
        Case "false": JSONValue = "false": Exit Function
    End Select

    Err.Clear
    On Error Resume Next
    d = CDbl(text)
    If Err.Number = 0 Then
        On Error GoTo 0
        If Not IsDate(text) Then
            ' 仅当字符串往返转换保持原始表示时才输出为数字
            ' (避免 "1000000000000000" → "1E+15")。
            ' Str$ 始终使用 "." 作为小数点，不受区域设置影响。
            If LTrim$(Str$(d)) = text Then
                JSONValue = LTrim$(Str$(d))
            Else
                ' 表示已改变 — 保留原始文本作为带引号的字符串
                JSONValue = """" & JSONEscape(text) & """"
            End If
            Exit Function
        End If
    Else
        Err.Clear
    End If
    On Error GoTo 0

    ' 日期字符串（涵盖所有常见格式，包括欧洲格式 dd.mm.yyyy）
    If IsDate(text) Then
        On Error Resume Next
        dt = CDate(text)
        If Err.Number = 0 Then
            On Error GoTo 0
            JSONValue = """" & Format$(dt, "yyyy-mm-dd""T""hh:nn:ss") & """"
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If

    JSONValue = """" & JSONEscape(text) & """"
End Function

'=============================================================================
' RangeToMarkdown — 区域转 Markdown 表格
'=============================================================================
Public Function RangeToMarkdown(ByRef rng As Range, _
    Optional ByVal includeHeader As Boolean = True) As String
    Dim data As Variant
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim parts() As String
    Dim pIdx As Long
    Dim cellText As String
    Dim startRow As Long

    If rng Is Nothing Then
        RangeToMarkdown = ""
        Exit Function
    End If

    GetRangeData rng, data, nRows, nCols

    pIdx = 0
    ReDim parts(0 To nRows * (nCols * 3 + 4) + 10)

    ' 表头
    If includeHeader And nRows >= 1 Then
        parts(pIdx) = "| ": pIdx = pIdx + 1
        For j = 1 To nCols
            If j > 1 Then
                parts(pIdx) = " | ": pIdx = pIdx + 1
            End If
            cellText = SafeText(data(1, j))
            cellText = Replace(cellText, "|", "&#124;")
            cellText = Replace(cellText, vbLf, "<br>")
            cellText = Replace(cellText, vbCr, "")
            parts(pIdx) = cellText: pIdx = pIdx + 1
        Next j
        parts(pIdx) = " |": pIdx = pIdx + 1
        parts(pIdx) = vbCrLf: pIdx = pIdx + 1

        ' 分隔行
        parts(pIdx) = "| ": pIdx = pIdx + 1
        For j = 1 To nCols
            parts(pIdx) = "---|": pIdx = pIdx + 1
        Next j
        parts(pIdx) = vbCrLf: pIdx = pIdx + 1
        startRow = 2
    Else
        startRow = 1
    End If

    ' 表体
    For i = startRow To nRows
        EnsureCapacity parts, pIdx, nCols * 3 + 2
        parts(pIdx) = "| ": pIdx = pIdx + 1
        For j = 1 To nCols
            If j > 1 Then
                parts(pIdx) = " | ": pIdx = pIdx + 1
            End If
            cellText = SafeText(data(i, j))
            cellText = Replace(cellText, "|", "&#124;")
            cellText = Replace(cellText, vbLf, "<br>")
            cellText = Replace(cellText, vbCr, "")
            parts(pIdx) = cellText: pIdx = pIdx + 1
        Next j
        parts(pIdx) = " |": pIdx = pIdx + 1
        parts(pIdx) = vbCrLf: pIdx = pIdx + 1
    Next i

    If pIdx > 0 Then
        ReDim Preserve parts(0 To pIdx - 1)
        RangeToMarkdown = Join(parts, "")
    End If
End Function

'=============================================================================
' CopyRangeToSheet — 将 Variant 数组快速写入工作表
'
' 自动根据数组大小选择输出范围
'=============================================================================
Public Sub CopyRangeToSheet( _
    ByRef data As Variant, _
    ByRef startCell As Range)

    Dim nRows As Long, nCols As Long
    Dim dest As Range, dest2 As Range, dest3 As Range
    Dim rowArr() As Variant, outArr() As Variant
    Dim j As Long, i As Long
    Dim r0 As Long, c0 As Long

    If startCell Is Nothing Then Exit Sub
    If IsEmpty(data) Then Exit Sub

    If Not IsArray(data) Then
        startCell.Value = data
        Exit Sub
    End If

    ' 检测空数组 / 未维度数组 (#74)
    Dim lbProbe As Long
    Err.Clear: On Error Resume Next
    lbProbe = LBound(data)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub  ' 未维度数组，静默退出
    End If
    On Error GoTo 0

    ' 检测 1D 还是 2D
    Err.Clear
    On Error Resume Next
    nCols = UBound(data, 2)
    If Err.Number <> 0 Then
        ' 1D 数组 → 单行输出
        On Error GoTo 0
        nRows = 1
        nCols = UBound(data) - LBound(data) + 1
        If nCols <= 0 Then Exit Sub
        If Not OutputFits(startCell, nRows, nCols) Then Exit Sub
        Set dest = startCell.Resize(nRows, nCols)
        If dest.HasFormula Or IsNull(dest.HasFormula) Then
            Err.Raise ERR_FORMULA_RANGE, "CopyRangeToSheet", _
                "目标区域包含公式，请先使用 ClearContents 清除公式后再复制。"
        End If
        ReDim rowArr(1 To nRows, 1 To nCols)
        For j = 1 To nCols
            rowArr(1, j) = data(LBound(data) + j - 1)
        Next j
        SafeWriteValue dest, rowArr, "CopyRangeToSheet"
        Exit Sub
    End If
    On Error GoTo 0

    nRows = UBound(data, 1) - LBound(data, 1) + 1
    nCols = UBound(data, 2) - LBound(data, 2) + 1
    If nRows <= 0 Or nCols <= 0 Then Exit Sub

    If LBound(data, 1) = 1 And LBound(data, 2) = 1 Then
        If OutputFits(startCell, nRows, nCols) Then
            Set dest2 = startCell.Resize(nRows, nCols)
            If dest2.HasFormula Or IsNull(dest2.HasFormula) Then
                Err.Raise ERR_FORMULA_RANGE, "CopyRangeToSheet", _
                    "目标区域包含公式，请先使用 ClearContents 清除公式后再复制。"
            End If
            SafeWriteValue dest2, data, "CopyRangeToSheet"
        End If
    Else
        ' 处理非 1-based 数组
        If Not OutputFits(startCell, nRows, nCols) Then Exit Sub
        Set dest3 = startCell.Resize(nRows, nCols)
        If dest3.HasFormula Or IsNull(dest3.HasFormula) Then
            Err.Raise ERR_FORMULA_RANGE, "CopyRangeToSheet", _
                "目标区域包含公式，请先使用 ClearContents 清除公式后再复制。"
        End If
        ReDim outArr(1 To nRows, 1 To nCols)
        r0 = LBound(data, 1)
        c0 = LBound(data, 2)
        For i = 1 To nRows
            For j = 1 To nCols
                outArr(i, j) = data(r0 + i - 1, c0 + j - 1)
            Next j
        Next i
        SafeWriteValue dest3, outArr, "CopyRangeToSheet"
    End If
End Sub

Private Function OutputFits(ByRef startCell As Range, ByVal nRows As Long, ByVal nCols As Long) As Boolean
    OutputFits = (startCell.Row + nRows - 1 <= startCell.Parent.Rows.Count) And _
                 (startCell.Column + nCols - 1 <= startCell.Parent.Columns.Count)
End Function

Private Sub SafeWriteValue(ByRef dest As Range, ByRef value As Variant, ByVal caller As String)
    If dest.MergeCells Then
        If dest.MergeArea.Address <> dest.Address Then
            Set dest = dest.MergeArea.Cells(1, 1)
        End If
    End If
    If dest.HasArray Then
        Err.Raise vbObjectError + 1005, caller, _
            "Target range contains an array formula. Clear the array formula before copying."
    End If
    dest.value = value
End Sub

'=============================================================================
' TransposeRange — 转置区域到目标位置
'=============================================================================
Public Sub TransposeRange(ByRef rng As Range, ByRef destCell As Range)
    If rng Is Nothing Or destCell Is Nothing Then Exit Sub

    Dim nRows As Long, nCols As Long
    If rng.Count = 1 Then
        destCell.Value = rng.Value
        Exit Sub
    End If

    nRows = rng.Rows.Count
    nCols = rng.Columns.Count

    Const MAX_TRANSPOSE As Long = 65536
    If nRows * nCols > MAX_TRANSPOSE Then Exit Sub
    If destCell.Row + nCols - 1 > destCell.Parent.Rows.Count Then Exit Sub
    If destCell.Column + nRows - 1 > destCell.Parent.Columns.Count Then Exit Sub

    ' 手动转置以保留 Empty 单元格（Application.Transpose 会将 Empty 转换为 0）
    Dim src As Variant, dst() As Variant
    If rng.Areas.Count > 1 Then
        Err.Raise ERR_MULTI_SELECT, "TransposeRange", "不支持多选区，请使用连续区域。"
    End If
    src = rng.Value
    ReDim dst(1 To nCols, 1 To nRows)
    Dim i As Long, j As Long
    For i = 1 To nRows
        For j = 1 To nCols
            dst(j, i) = src(i, j)
        Next j
    Next i
    Dim dest As Range: Set dest = destCell.Resize(nCols, nRows)
    If dest.HasFormula Or IsNull(dest.HasFormula) Then
        Err.Raise ERR_FORMULA_RANGE, "TransposeRange", _
            "目标区域包含公式，请先使用 ClearContents 清除公式后再转置。"
    End If
    dest.Value = dst
End Sub

'=============================================================================
' ClearRangeContents — 清除区域内容 (保留格式)
'=============================================================================
Public Sub ClearRangeContents(ByRef rng As Range)
    If Not rng Is Nothing Then rng.ClearContents
End Sub

'=============================================================================
' ClearRangeFormats — 清除区域格式 (保留内容)
'=============================================================================
Public Sub ClearRangeFormats(ByRef rng As Range)
    If Not rng Is Nothing Then rng.ClearFormats
End Sub

'=============================================================================
' AutoFitRange — 自动调整列宽/行高
'
' 参数:
'   rng      - 目标区域
'   fitCols  - 调整列宽 (默认 True)
'   fitRows  - 调整行高 (默认 True)
'=============================================================================
Public Sub AutoFitRange(ByRef rng As Range, _
    Optional ByVal fitCols As Boolean = True, _
    Optional ByVal fitRows As Boolean = True)
    If rng Is Nothing Then Exit Sub
    If fitCols Then rng.Columns.AutoFit
    If fitRows Then rng.Rows.AutoFit
End Sub


'=============================================================================
' RemoveDuplicatesRange — 按指定列去重
'=============================================================================
Public Sub RemoveDuplicatesRange( _
    ByRef rng As Range, _
    Optional ByVal colIndices As Variant, _
    Optional ByVal hasHeader As XlYesNoGuess = xlYes)

    If rng Is Nothing Then Exit Sub

    If IsMissing(colIndices) Then
        rng.RemoveDuplicates Columns:=1, Header:=hasHeader
    Else
        If IsArray(colIndices) Then
            rng.RemoveDuplicates Columns:=colIndices, Header:=hasHeader
        Else
            rng.RemoveDuplicates Columns:=Array(CLng(colIndices)), Header:=hasHeader
        End If
    End If
End Sub

'=============================================================================
' SortRange — 按指定列排序
'=============================================================================
Public Sub SortRange( _
    ByRef rng As Range, _
    ByVal keyCol As Long, _
    Optional ByVal ascending As Boolean = True, _
    Optional ByVal hasHeader As Boolean = True)

    If rng Is Nothing Then Exit Sub
    If keyCol < 1 Or keyCol > rng.Columns.Count Then Exit Sub

    Dim order As XlSortOrder
    If ascending Then order = xlAscending Else order = xlDescending

    If hasHeader Then
        rng.Sort Key1:=rng.Columns(keyCol), Order1:=order, Header:=xlYes
    Else
        rng.Sort Key1:=rng.Columns(keyCol), Order1:=order, Header:=xlNo
    End If
End Sub

'=============================================================================
' FilterRangeToArray — 条件筛选返回二维 Variant 数组
'
' 参数:
'   rng    - 源区域 (含标题行)
'   colIdx - 筛选列 (1-based)
'   op     - 运算符: "=" "<>" "<" ">" "<=" ">=" "contains" "notcontains" "startswith" "endswith" "isblank" "isnotblank"
'   value  - 比较值
'   includeHeader - 第一行是否为标题行 (默认 True)
'
' 返回: 符合条件的行组成的 2D 数组
'=============================================================================
Public Function FilterRangeToArray( _
    ByRef rng As Range, _
    ByVal colIdx As Long, _
    ByVal op As String, _
    ByVal value As Variant, _
    Optional ByVal includeHeader As Boolean = True) As Variant

    Dim data As Variant
    Dim result() As Variant
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long, outIdx As Long

    If rng Is Nothing Then
        FilterRangeToArray = Array()
        Exit Function
    End If

    GetRangeData rng, data, nRows, nCols

    If colIdx < 1 Or colIdx > nCols Then
        FilterRangeToArray = Array()
        Exit Function
    End If

    ' 第一遍: 计数
    Dim cnt As Long: cnt = 0
    Dim startRow As Long
    If includeHeader Then
        cnt = 1 ' 标题行
        startRow = 2
    Else
        startRow = 1
    End If
    For i = startRow To nRows
        If FilterPasses(data(i, colIdx), value, op) Then
            cnt = cnt + 1
        End If
    Next i

    If cnt = 0 Then
        FilterRangeToArray = Array()
        Exit Function
    End If

    ReDim result(1 To cnt, 1 To nCols)
    ' 复制标题行
    If includeHeader Then
        For j = 1 To nCols
            result(1, j) = data(1, j)
        Next j
    End If

    If includeHeader Then outIdx = 2 Else outIdx = 1
    For i = startRow To nRows
        If FilterPasses(data(i, colIdx), value, op) Then
            For j = 1 To nCols
                result(outIdx, j) = data(i, j)
            Next j
            outIdx = outIdx + 1
        End If
    Next i
    FilterRangeToArray = result
End Function

' FilterPasses — delegates to VariantKit.FilterPasses (single source of truth)
Private Function FilterPasses(ByVal cellVal As Variant, _
    ByVal criteria As Variant, ByVal op As String) As Boolean
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    FilterPasses = vk.FilterPasses(cellVal, criteria, op)
End Function

' 类型感知的相等比较，避免与区域设置相关的 CStr 比较。
' 委托给 VariantKit.ValuesEqual (单一数据源, 含防御性 CDate 守卫)
Public Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    ValuesEqual = vk.ValuesEqual(a, b)
End Function


'=============================================================================
' RangeToArray — 区域→二维 Variant 数组 (无筛选，直接转换)
'
' 返回: 1-based 二维 Variant 数组
'=============================================================================
Public Function RangeToArray(ByRef rng As Range) As Variant()
    Dim result() As Variant
    If rng Is Nothing Then
        ReDim result(1 To 1, 1 To 1): Erase result
        RangeToArray = result
        Exit Function
    End If

    If rng.Areas.Count > 1 Then
        Err.Raise ERR_MULTI_SELECT, "RangeToArray", "不支持多选区，请使用连续区域。"
    End If
    If rng.Count = 1 Then
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = rng.Value
    Else
        result = rng.Value
    End If
    RangeToArray = result
End Function


'=============================================================================
' RangeExists — 检查区域引用是否有效 (不是 #REF!)
'=============================================================================
Public Function RangeExists(ByRef rng As Range) As Boolean
    If rng Is Nothing Then
        RangeExists = False
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    Dim addr As String
    addr = rng.Address
    RangeExists = (Err.Number = 0)
    On Error GoTo 0
End Function

'=============================================================================
' NamedRangeExists — 检查名称是否存在
'=============================================================================
Public Function NamedRangeExists(ByVal rangeName As String, _
                                  Optional ByRef wb As Workbook = Nothing) As Boolean

    If Len(rangeName) = 0 Then
        NamedRangeExists = False
        Exit Function
    End If

    If wb Is Nothing Then Set wb = Application.ActiveWorkbook

    On Error Resume Next
    Dim rng As Range
    Set rng = wb.Names(rangeName).RefersToRange
    NamedRangeExists = (Err.Number = 0)
    On Error GoTo 0
End Function

'=============================================================================
' NamedRangeAdd — 创建或更新命名区域
'
' 参数:
'   rangeName - 名称
'   rng    - 目标区域
'   wb     - 工作簿 (默认 ActiveWorkbook)
'=============================================================================
Public Sub NamedRangeAdd(ByVal rangeName As String, ByRef rng As Range, _
                          Optional ByRef wb As Workbook = Nothing)
    Dim savedErr As Long, savedDesc As String
    Dim tmp As Name

    If Len(rangeName) = 0 Then Exit Sub
    If rng Is Nothing Then Exit Sub
    If wb Is Nothing Then Set wb = Application.ActiveWorkbook

    ' 更新或插入：先删除已有名称（Names.Add 对已存在名称会抛出错误）
    On Error Resume Next
    wb.Names(rangeName).Delete
    If Err.Number <> 0 Then
        savedErr = Err.Number
        savedDesc = Err.Description
        On Error GoTo 0
        ' 如果删除失败是因为名称不存在，则没问题。
        ' 如果删除失败是其他原因（工作簿受保护等），则重新抛出错误。
        On Error Resume Next
        Set tmp = wb.Names(rangeName)
        If Err.Number = 0 Then
            On Error GoTo 0
            Err.Raise savedErr, "NamedRangeAdd", "Cannot update '" & rangeName & "': " & savedDesc
        End If
        On Error GoTo 0
    End If
    On Error GoTo 0
    wb.Names.Add Name:=rangeName, RefersTo:=rng
End Sub

'=============================================================================
' NamedRangeDelete — 删除命名区域
'
' 参数:
'   name - 名称
'   wb   - 工作簿 (默认 ActiveWorkbook)
'=============================================================================
Public Sub NamedRangeDelete(ByVal rangeName As String, _
                             Optional ByRef wb As Workbook = Nothing)
    If Len(rangeName) = 0 Then Exit Sub
    If wb Is Nothing Then Set wb = Application.ActiveWorkbook

    On Error Resume Next
    wb.Names(rangeName).Delete
    On Error GoTo 0
End Sub


'=============================================================================
' IsRangeEmpty — 检查区域是否有数据
'=============================================================================
Public Function IsRangeEmpty(ByRef rng As Range) As Boolean
    If rng Is Nothing Then
        IsRangeEmpty = True
        Exit Function
    End If
    IsRangeEmpty = (Application.CountA(rng) = 0)
End Function

'=============================================================================
' CountVisible — 统计可见 (非隐藏) 行数
'=============================================================================
Public Function CountVisible(ByRef rng As Range) As Long
    If rng Is Nothing Then Exit Function
    Dim area As Range, visCells As Range
    ' 使用 SpecialCells(xlCellTypeVisible) 实现每个 Area O(1) 次 COM 调用
    ' 避免逐行的 O(n) Hidden 检查
    For Each area In rng.Areas
        Err.Clear
        On Error Resume Next
        Set visCells = area.Columns(1).SpecialCells(xlCellTypeVisible)
        If Err.Number = 0 Then CountVisible = CountVisible + visCells.Count
        Err.Clear
    Next area
    On Error GoTo 0
End Function

'=============================================================================
' RangeDiff — 比较两个同样大小的区域
'
' 返回: 差异单元格的 Range, 或 Nothing (无差异)
' 差异数超过 ~8,000 时 outTruncated = True（Excel Union 上限）
'=============================================================================
Public Function RangeDiff(ByRef rng1 As Range, ByRef rng2 As Range, _
                          Optional ByRef outTruncated As Boolean) As Range
    Dim result As Range
    Dim cell As Range
    Dim nRows As Long, nCols As Long
    Dim i As Long, j As Long
    Dim data1 As Variant, data2 As Variant

    outTruncated = False

    If rng1 Is Nothing Or rng2 Is Nothing Then
        Set RangeDiff = Nothing
        Exit Function
    End If

    If rng1.Areas.Count > 1 Or rng2.Areas.Count > 1 Then
        Err.Raise ERR_MULTI_AREA, "RangeDiff","Multi-area ranges are not supported"
    End If

    nRows = rng1.Rows.Count
    nCols = rng1.Columns.Count

    If rng2.Rows.Count <> nRows Or rng2.Columns.Count <> nCols Then
        Err.Raise ERR_MULTI_SELECT, "RangeDiff","Ranges must have the same dimensions"
    End If

    If rng1.Count = 1 Then
        If Not ValuesEqual(rng1.Value, rng2.Value) Then
            Set RangeDiff = rng1
        End If
        Exit Function
    End If

    data1 = rng1.Value: data2 = rng2.Value
    Dim first As Boolean: first = True
    Const MAX_UNION_AREAS As Long = 8000

    For i = 1 To nRows
        For j = 1 To nCols
            If Not ValuesEqual(data1(i, j), data2(i, j)) Then
                Set cell = rng1.Cells(i, j)
                If first Then
                    Set result = cell
                    first = False
                Else
                    If result.Areas.Count < MAX_UNION_AREAS Then
                        Set result = Application.Union(result, cell)
                    Else
                        outTruncated = True
                        Exit For
                    End If
                End If
            End If
        Next j
        If outTruncated Then Exit For
    Next i

    Set RangeDiff = result
End Function



'=============================================================================
' GetCellAddress — 构建单元格地址字符串
'
' 例: GetCellAddress(3, 5) → "$E$3"
'=============================================================================
Public Function GetCellAddress( _
    ByVal row As Long, _
    ByVal col As Long, _
    Optional ByVal absolute As Boolean = True) As String

    If row < 1 Or col < 1 Then
        GetCellAddress = ""
        Exit Function
    End If

    If absolute Then
        GetCellAddress = "$" & ColLetter(col) & "$" & row
    Else
        GetCellAddress = ColLetter(col) & row
    End If
End Function

'=============================================================================
' 工作表函数 (UDF_RANGE_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_RANGE_LASTROW(Optional ByVal ws As Variant = Nothing) As Variant
    On Error GoTo EH: UDF_RANGE_LASTROW = LastRow(ws): Exit Function
EH: UDF_RANGE_LASTROW = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_LASTCOL(Optional ByVal ws As Variant = Nothing) As Variant
    On Error GoTo EH: UDF_RANGE_LASTCOL = LastCol(ws): Exit Function
EH: UDF_RANGE_LASTCOL = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_FIRSTROW(Optional ByVal ws As Variant = Nothing) As Variant
    On Error GoTo EH: UDF_RANGE_FIRSTROW = FirstRow(ws): Exit Function
EH: UDF_RANGE_FIRSTROW = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_FIRSTCOL(Optional ByVal ws As Variant = Nothing) As Variant
    On Error GoTo EH: UDF_RANGE_FIRSTCOL = FirstCol(ws): Exit Function
EH: UDF_RANGE_FIRSTCOL = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_COL_LETTER(ByVal colNum As Variant) As Variant
    On Error GoTo EH: UDF_RANGE_COL_LETTER = ColLetter(colNum): Exit Function
EH: UDF_RANGE_COL_LETTER = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_COL_NUM(ByVal colRef As Variant) As Variant
    On Error GoTo EH
    Dim r As Long: r = ColNumber(colRef)
    If r = -1 Then UDF_RANGE_COL_NUM = CVErr(xlErrValue) Else UDF_RANGE_COL_NUM = r
    Exit Function
EH: UDF_RANGE_COL_NUM = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_EXISTS(ByVal rng As Variant) As Variant
    On Error GoTo EH
    If Not TypeOf rng Is Range Then UDF_RANGE_EXISTS = CVErr(xlErrValue): Exit Function
    Dim reRng As Range: Set reRng = rng
    UDF_RANGE_EXISTS = RangeExists(reRng)
    Exit Function
EH: UDF_RANGE_EXISTS = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_NAMEDEXISTS(ByVal rangeName As Variant, Optional ByVal wb As Variant = Nothing) As Variant
    On Error GoTo EH: UDF_RANGE_NAMEDEXISTS = NamedRangeExists(rangeName, wb): Exit Function
EH: UDF_RANGE_NAMEDEXISTS = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_ISEMPTY(ByVal rng As Variant) As Variant
    On Error GoTo EH
    If Not TypeOf rng Is Range Then UDF_RANGE_ISEMPTY = CVErr(xlErrValue): Exit Function
    Dim ieRng As Range: Set ieRng = rng
    UDF_RANGE_ISEMPTY = IsRangeEmpty(ieRng)
    Exit Function
EH: UDF_RANGE_ISEMPTY = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_COUNTVISIBLE(ByVal rng As Variant) As Variant
    On Error GoTo EH
    If Not TypeOf rng Is Range Then UDF_RANGE_COUNTVISIBLE = CVErr(xlErrValue): Exit Function
    Dim cvRng As Range: Set cvRng = rng
    UDF_RANGE_COUNTVISIBLE = CountVisible(cvRng)
    Exit Function
EH: UDF_RANGE_COUNTVISIBLE = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_CELLADDRESS(ByVal row As Variant, ByVal col As Variant, _
    Optional ByVal absolute As Variant = True) As Variant
    On Error GoTo EH: UDF_RANGE_CELLADDRESS = GetCellAddress(row, col, absolute): Exit Function
EH: UDF_RANGE_CELLADDRESS = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_SAFETEXT(ByVal v As Variant) As Variant
    On Error GoTo EH:     UDF_RANGE_SAFETEXT = SafeText(v): Exit Function
EH: UDF_RANGE_SAFETEXT = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_TOHTML(ByVal rng As Variant, _
    Optional ByVal includeHeader As Variant = True, Optional ByVal tableClass As Variant = "") As Variant
    On Error GoTo EH
    If Not TypeOf rng Is Range Then UDF_RANGE_TOHTML = CVErr(xlErrValue): Exit Function
    Dim thRng As Range: Set thRng = rng
    UDF_RANGE_TOHTML = RangeToHTML(thRng, includeHeader, tableClass)
    Exit Function
EH: UDF_RANGE_TOHTML = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_TOJSON(ByVal rng As Variant, Optional ByVal includeHeader As Variant = True) As Variant
    On Error GoTo EH
    If Not TypeOf rng Is Range Then UDF_RANGE_TOJSON = CVErr(xlErrValue): Exit Function
    Dim tjRng As Range: Set tjRng = rng
    UDF_RANGE_TOJSON = RangeToJSON(tjRng, includeHeader)
    Exit Function
EH: UDF_RANGE_TOJSON = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_TOMD(ByVal rng As Variant, Optional ByVal includeHeader As Variant = True) As Variant
    On Error GoTo EH
    If Not TypeOf rng Is Range Then UDF_RANGE_TOMD = CVErr(xlErrValue): Exit Function
    Dim tmRng As Range: Set tmRng = rng
    UDF_RANGE_TOMD = RangeToMarkdown(tmRng, includeHeader)
    Exit Function
EH: UDF_RANGE_TOMD = CVErr(xlErrValue)
End Function

Public Function UDF_RANGE_FILTER(ByVal rng As Variant, ByVal colIdx As Variant, _
    ByVal op As Variant, ByVal value As Variant, Optional ByVal includeHeader As Variant = True) As Variant
    On Error GoTo EH
    If Not TypeOf rng Is Range Then UDF_RANGE_FILTER = CVErr(xlErrValue): Exit Function
    Dim ftRng As Range: Set ftRng = rng
    UDF_RANGE_FILTER = FilterRangeToArray(ftRng, colIdx, op, value, includeHeader)
    Exit Function
EH: UDF_RANGE_FILTER = CVErr(xlErrValue)
End Function

'=====================================================================
' 使用示例
'=====================================================================
' LastRow()                                              → 当前工作表最后有数据的行号
' LastCol()                                              → 当前工作表最后有数据的列号
' ColLetter(27)                                          → "AA"
' ColNumber("AA")                                        → 27
' UsedRangeEx().Address                                  → "$A$1:$F$100"
' RangeToHTML(Range("A1:B3"), , "mytable")               → HTML 表格
' ExportRangeToCSV Range("A1:C10"), "C:\data.csv"        → 导出 CSV

'=====================================================================