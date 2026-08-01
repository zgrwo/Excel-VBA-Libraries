Option Explicit

'==============================================================================
' Module:       PivotUtils
' Purpose:      Data reshaping: pivot, unpivot, group-by, cross-join
' Layer:        Data
' Dependencies: VBA-Core (VariantKit, ArrayOps, DictProxy)
' Public:       15 functions/subs
'==============================================================================

Private DP As New DictProxy

'=====================================================================
' PivotUtils.bas — 数据重塑（透视/逆透视/分组/过滤/交叉连接）
'
' 注意: 本模块函数不支持多区域 Range (Multi-Area)，传入将导致运行时错误 1004。
'
' 功能: 将源数据表按指定行/列维度进行透视转换，值按出现顺序展开到字母后缀列
'
' 工作表函数 (UDF_PIVOT_*):
'   UDF_PIVOT_CONVERT     — 数据透视
'   UDF_PIVOT_GROUPBY   — 分组聚合 (SUM/COUNT/AVG/MIN/MAX)
'   UDF_PIVOT_VLOOKUP   — 数组版 VLOOKUP
'   UDF_PIVOT_CROSSJOIN — 交叉连接 (笛卡尔积)
'
' 透视/逆透视:
'   RawConversion          — 工作表函数 (接受 Range)
'   RawConversionFromArray — VBA 函数 (接受二维数组)
' 调用约定 — 参数顺序: (valueColRef, colDimRef, rowDimRef)
'   valueColRef — 值填充列（其值填入交叉表单元格）
'   colDimRef  — 列维度列（其唯一值展开为列标题）
'   rowDimRef  — 行维度列（其唯一值作为输出行标签）
'
'   RawConversionToRange   — 透视结果输出到工作表
'   Unpivot                — 逆透视 (宽表 → 长表)
'
' 分组/筛选:
'   GroupBy                — 分组聚合 (SUM/COUNT/AVG/MIN/MAX)
'   FilterTable            — 按条件筛选表格
'
' 操作:
'   SplitColumnToRows      — 拆分列到多行
'   MergeColumns           — 合并多列
'   CrossJoin              — 交叉连接
'   TransposeTable         — 表格转置
'
' 查找/辅助:
'   VLookupArray           — 数组版 VLOOKUP
'=====================================================================

Private Const ALPHABET As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const MAX_EXCEL_ROWS As Long = 1048576
Private Const MAX_EXCEL_COLS As Long = 16384
Private Const MAX_SORT_DEPTH As Long = 5000

' -- 错误代码常量 -----------------------------------------------------
Private Const ERR_SOURCE         As String = "PivotUtils"
Private Const ERR_INVALID_INPUT  As Long = vbObjectError + 1001
Private Const ERR_OUT_OF_BOUNDS  As Long = vbObjectError + 1002
Private Const ERR_DUPLICATE_NAME As Long = vbObjectError + 1030
Private Const ERR_NAME_NOT_FOUND As Long = vbObjectError + 1031
Private Const ERR_EMPTY_DIM      As Long = vbObjectError + 1040
Private Const ERR_SORT_OVERFLOW  As Long = vbObjectError + 1051
Private Const ERR_PIVOT_NO_ID    As Long = vbObjectError + 1082
Private Const ERR_PIVOT_NO_VAL   As Long = vbObjectError + 1083
Private Const ERR_PIVOT_OVERFLOW As Long = vbObjectError + 1085
Private Const ERR_MULTI_AREA       As Long = vbObjectError + 1003
Private Const ERR_PARAM_TYPE     As Long = vbObjectError + 1010
Private Const ERR_PARAM_MULTI    As Long = vbObjectError + 1011
Private Const ERR_PARAM_NOT_STR  As Long = vbObjectError + 1012
Private Const ERR_PARAM_EMPTY    As Long = vbObjectError + 1013
Private Const ERR_CELL_EMPTY     As Long = vbObjectError + 1020
Private Const ERR_RESOLVE_EMPTY  As Long = vbObjectError + 1021
Private Const ERR_EMPTY_DIM2     As Long = vbObjectError + 1041
Private Const ERR_SUFFIX_RANGE   As Long = vbObjectError + 1050
Private Const ERR_FILTER_SRC     As Long = vbObjectError + 1120
Private Const ERR_FILTER_DEST    As Long = vbObjectError + 1121
Private Const ERR_FILTER_OP      As Long = vbObjectError + 1062
Private Const ERR_UNPIVOT_SRC    As Long = vbObjectError + 1122
Private Const ERR_UNPIVOT_DEST   As Long = vbObjectError + 1123
Private Const ERR_UNPIVOT_COLS   As Long = vbObjectError + 1084
Private Const ERR_SPLIT_SRC      As Long = vbObjectError + 1090
Private Const ERR_SPLIT_DEST     As Long = vbObjectError + 1091
Private Const ERR_SPLIT_DELIM    As Long = vbObjectError + 1092
Private Const ERR_MERGE_SRC      As Long = vbObjectError + 1124
Private Const ERR_MERGE_DEST     As Long = vbObjectError + 1125
Private Const ERR_TRANSTABLE_SRC As Long = vbObjectError + 1110
Private Const ERR_TRANSTABLE_DST As Long = vbObjectError + 1111


'=============================================================================
' 工作表函数: RawConversion (接受 Range)
'=============================================================================
Public Function RawConversion( _
    ByVal srcRange As Range, _
    ByVal valueColRef As Variant, _
    ByVal colDimRef As Variant, _
    ByVal rowDimRef As Variant, _
    Optional ByVal keepBlank As Boolean = True, _
    Optional ByVal sortLabels As Boolean = False _
) As Variant

    On Error GoTo ErrHandler

    Dim valueColName As String
    Dim colDimName As String
    Dim rowDimName As String

    ValidateInput srcRange, valueColRef, colDimRef, rowDimRef

    valueColName = ResolveHeaderText(valueColRef, "value")
    colDimName = ResolveHeaderText(colDimRef, "column")
    rowDimName = ResolveHeaderText(rowDimRef, "row")

    RawConversion = RawConversionFromArray(srcRange.Value, valueColName, colDimName, rowDimName, keepBlank, sortLabels)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "RawConversion", Err.Description
End Function

'=============================================================================
' VBA 函数: RawConversionFromArray (接受二维数组)
'=============================================================================
Public Function RawConversionFromArray( _
    ByVal dataArray As Variant, _
    ByVal valueColName As String, _
    ByVal colDimName As String, _
    ByVal rowDimName As String, _
    Optional ByVal keepBlank As Boolean = True, _
    Optional ByVal sortLabels As Boolean = False _
) As Variant

    If Not IsArray(dataArray) Then
        Err.Raise ERR_INVALID_INPUT, "RawConversionFromArray", "输入必须是 2D 数组。"
    End If

    Dim nTotalRows As Long, nCols As Long
    Dim r0 As Long, c0 As Long
    r0 = LBound(dataArray, 1)
    c0 = LBound(dataArray, 2)
    nTotalRows = UBound(dataArray, 1) - r0 + 1
    nCols = UBound(dataArray, 2) - c0 + 1

    If nTotalRows < 2 Then
        Err.Raise ERR_OUT_OF_BOUNDS, "RawConversionFromArray", _
            "输入数组至少需要包含 1 行标题 + 1 行数据 (>=2 行), 实际为 " & nTotalRows & " 行。"
    End If

    ' 提取标题行 (1-based 便于 FindColumnIndex)
    Dim headerRow() As Variant
    ReDim headerRow(1 To nCols)
    Dim c As Long
    For c = 1 To nCols
        headerRow(c) = dataArray(r0, c0 + c - 1)
    Next c

    Dim nDataRows As Long
    nDataRows = nTotalRows - 1

    ' 定位三个关键列的索引 (1-based in headerRow)
    Dim idxValueCol As Long
    Dim idxColDim As Long
    Dim idxRowDim As Long
    idxValueCol = FindColumnIndex(headerRow, valueColName, "valueCol: """ & valueColName & """")
    idxColDim = FindColumnIndex(headerRow, colDimName, "colDim: """ & colDimName & """")
    idxRowDim = FindColumnIndex(headerRow, rowDimName, "rowDim: """ & rowDimName & """")

    ' Map from headerRow (1-based) back to dataArray column index
    ' e.g. c0=0 → dataArray(dr, 0) = first column, c0=1 → dataArray(dr, 1) = first col
    idxValueCol = c0 + idxValueCol - 1
    idxColDim = c0 + idxColDim - 1
    idxRowDim = c0 + idxRowDim - 1

    ' 构建分组字典 (直接索引 dataArray，避免完整复制数据区域)
    Dim groups As Object
    Dim colValues As Object
    Dim rowValues As Object
    Dim colMaxItems As Object
    BuildGroups dataArray, nDataRows, r0 + 1, idxRowDim, idxColDim, idxValueCol, _
                groups, colValues, rowValues, colMaxItems

    Dim result() As Variant
    BuildResultArray groups, colValues, rowValues, colMaxItems, _
                     keepBlank, sortLabels, CStr(headerRow(idxRowDim - c0 + 1)), result
    RawConversionFromArray = result
End Function

'=============================================================================
' 辅助: 验证输入参数
'=============================================================================
Private Sub ValidateInput( _
    ByVal srcRange As Range, _
    ByVal valueColRef As Variant, _
    ByVal colDimRef As Variant, _
    ByVal rowDimRef As Variant _
)
    If srcRange Is Nothing Then
        Err.Raise ERR_INVALID_INPUT, "RawConversion", "源数据区域不能为空。"
    End If

    If srcRange.Areas.Count > 1 Then
        Err.Raise ERR_OUT_OF_BOUNDS, "RawConversion", "源数据区域必须为连续区域，不支持多选区。"
    End If

    If srcRange.Rows.Count < 2 Then
        Err.Raise ERR_EMPTY_DIM, "RawConversion", _
            "源数据区域至少需要包含 1 行标题 + 1 行数据 (>=2 行), 当前为 " & srcRange.Rows.Count & " 行。"
    End If

    ValidateSingleParam valueColRef, "value"
    ValidateSingleParam colDimRef, "column"
    ValidateSingleParam rowDimRef, "row"
End Sub

'=============================================================================
' 辅助: 校验单个参数 (支持 Range 或 String)
'=============================================================================
Private Sub ValidateSingleParam(ByVal param As Variant, ByVal paramName As String)
    Dim rng As Range
    If IsObject(param) Then
        If TypeOf param Is Range Then
            Set rng = param
            If rng.Cells.Count > 1 Then
                Err.Raise ERR_PARAM_MULTI, "RawConversion", _
                    "参数 """ & paramName & """ 引用了多个单元格 (" & rng.Address(False, False) & ")，请传入单个单元格。"
            End If
        Else
            Err.Raise ERR_PARAM_TYPE, "RawConversion", _
                "参数 """ & paramName & """ 是对象但并非单元格区域 (类型: " & TypeName(param) & ")，请传入单元格引用或纯文本。"
        End If
    Else
        If VarType(param) <> vbString Then
            Err.Raise ERR_PARAM_NOT_STR, "RawConversion", _
                "参数 """ & paramName & """ 不是字符串也不是单元格引用 (类型: " & TypeName(param) & ")，请传入单元格或文本。"
        End If
        If Len(Trim(CStr(param))) = 0 Then
            Err.Raise ERR_PARAM_EMPTY, "RawConversion", _
                "参数 """ & paramName & """ 是空字符串，请提供有效的标题文本。"
        End If
    End If
End Sub

'=============================================================================
' 辅助: 将参数解析为标题文本
'=============================================================================
Private Function ResolveHeaderText(ByVal param As Variant, ByVal paramName As String) As String
    Dim rng As Range
    Dim cellVal As Variant
    If IsObject(param) Then
        Set rng = param
        cellVal = rng.Value

        If IsEmpty(cellVal) Then
            Err.Raise ERR_CELL_EMPTY, "RawConversion", _
                "参数 """ & paramName & """ 引用的单元格 " & rng.Address(False, False) & " 为空，请填入标题文本。"
        End If

        ResolveHeaderText = Trim(CStr(cellVal))
    Else
        ResolveHeaderText = Trim(CStr(param))
    End If

    If Len(ResolveHeaderText) = 0 Then
        Err.Raise ERR_RESOLVE_EMPTY, "RawConversion", _
            "参数 """ & paramName & """ 解析后的标题文本为空。"
    End If
End Function

'=============================================================================
' 辅助: 在标题行中查找匹配的列索引
'=============================================================================
Private Function FindColumnIndex( _
    ByRef headerRow As Variant, _
    ByVal targetText As String, _
    ByVal sourceDescription As String _
) As Long
    Dim c As Long
    Dim trimmedTarget As String
    trimmedTarget = Trim(targetText)

    Dim foundIndex As Long
    foundIndex = -1

    For c = LBound(headerRow) To UBound(headerRow)
        If StrComp(Trim(CStr(headerRow(c))), trimmedTarget, vbTextCompare) = 0 Then
            If foundIndex = -1 Then
                foundIndex = c
            Else
                Err.Raise ERR_DUPLICATE_NAME, "RawConversion", _
                    "标题行中存在重复的列名 """ & trimmedTarget & """ (" & sourceDescription & ")。" & vbCrLf & _
                    "重复位置: 第 " & foundIndex & " 列 与 第 " & c & " 列。"
            End If
        End If
    Next c

    If foundIndex = -1 Then
        Dim headers As String
        headers = JoinHeaderNames(headerRow)
        Err.Raise ERR_NAME_NOT_FOUND, "RawConversion", _
            "在标题行中未找到 """ & trimmedTarget & """ (" & sourceDescription & ")。" & vbCrLf & _
            "可用的标题: " & headers
    End If

    FindColumnIndex = foundIndex
End Function

'=============================================================================
' 辅助: 将标题行拼接为可读字符串
'=============================================================================
Private Function JoinHeaderNames(ByRef headerRow As Variant) As String
    Dim c As Long
    Dim parts() As String
    ReDim parts(LBound(headerRow) To UBound(headerRow))

    For c = LBound(headerRow) To UBound(headerRow)
        parts(c) = """" & CStr(headerRow(c)) & """"
    Next c
    JoinHeaderNames = Join(parts, ", ")
End Function

'=============================================================================
' 辅助: 构建分组字典
'=============================================================================
Private Sub BuildGroups( _
    ByRef dataArray As Variant, _
    ByVal nDataRows As Long, _
    ByVal rowOffset As Long, _
    ByVal idxRowDim As Long, _
    ByVal idxColDim As Long, _
    ByVal idxValueCol As Long, _
    ByRef groups As Object, _
    ByRef colValues As Object, _
    ByRef rowValues As Object, _
    ByRef colMaxItems As Object _
)
    Set groups = DP.Create()
    Set colValues = DP.Create()
    Set rowValues = DP.Create()
    Set colMaxItems = DP.Create()

    Dim r As Long, dr As Long
    Dim rowDimVal As String
    Dim colDimVal As String
    Dim cellValue As Variant
    Dim innerDict As Object
    Dim arrList As Object

    For r = 1 To nDataRows
        dr = r + rowOffset - 1
        ' 跳过含错误值的关键列
        If IsError(dataArray(dr, idxRowDim)) Or IsNull(dataArray(dr, idxRowDim)) _
            Or IsError(dataArray(dr, idxColDim)) Or IsNull(dataArray(dr, idxColDim)) _
            Or IsError(dataArray(dr, idxValueCol)) Or IsNull(dataArray(dr, idxValueCol)) Then
            GoTo NextRow
        End If

        rowDimVal = Trim(CStr(dataArray(dr, idxRowDim)))
        If Len(rowDimVal) = 0 Then GoTo NextRow

        colDimVal = Trim(CStr(dataArray(dr, idxColDim)))
        If Len(colDimVal) = 0 Then GoTo NextRow

        cellValue = dataArray(dr, idxValueCol)

        If Not rowValues.Exists(rowDimVal) Then
            rowValues.Add rowDimVal, True
        End If

        If Not colValues.Exists(colDimVal) Then
            colValues.Add colDimVal, True
        End If

        If Not groups.Exists(rowDimVal) Then
            groups.Add rowDimVal, DP.Create()
        End If

        Set innerDict = groups(rowDimVal)

        If Not innerDict.Exists(colDimVal) Then
            innerDict.Add colDimVal, New Collection
        End If

        Set arrList = innerDict(colDimVal)
        arrList.Add cellValue

        If Not colMaxItems.Exists(colDimVal) Then
            colMaxItems.Add colDimVal, arrList.Count
        ElseIf arrList.Count > CLng(colMaxItems(colDimVal)) Then
            colMaxItems(colDimVal) = arrList.Count
        End If

NextRow:
    Next r

    If rowValues.Count = 0 Then
        Err.Raise ERR_EMPTY_DIM, "RawConversion", _
            "行标签列 (第" & idxRowDim & "列) 的值均为空，无法生成输出。"
    End If
    If colValues.Count = 0 Then
        Err.Raise ERR_EMPTY_DIM2, "RawConversion", _
            "列分组列 (第" & idxColDim & "列) 的值均为空，无法生成输出。"
    End If
End Sub

'=============================================================================
' 辅助: 构建结果数组
'=============================================================================
Private Sub BuildResultArray( _
    ByRef groups As Object, _
    ByRef colValues As Object, _
    ByRef rowValues As Object, _
    ByRef colMaxItems As Object, _
    ByVal keepBlank As Boolean, _
    ByVal sortLabels As Boolean, _
    ByVal rowDimHeader As Variant, _
    ByRef result As Variant _
)
    Dim colList() As String
    Dim rowList() As String

    CollectKeys colValues, colList
    CollectKeys rowValues, rowList

    If sortLabels Then
        SortStringArray colList
        SortStringArray rowList
    End If

    Dim nUniqueCols As Long, nUniqueRows As Long
    nUniqueCols = UBound(colList) - LBound(colList) + 1
    nUniqueRows = UBound(rowList) - LBound(rowList) + 1

    ' 计算总列数
    Dim nColsOut As Long, ci As Long
    nColsOut = 1
    For ci = LBound(colList) To UBound(colList)
        nColsOut = nColsOut + CLng(colMaxItems(colList(ci)))
    Next ci
    Dim nRowsOut As Long
    nRowsOut = 1 + nUniqueRows

    ReDim result(1 To nRowsOut, 1 To nColsOut)

    ' 填充标题行
    result(1, 1) = rowDimHeader

    Dim li As Long, colWidth As Long, outCol As Long
    outCol = 1
    For ci = LBound(colList) To UBound(colList)
        colWidth = CLng(colMaxItems(colList(ci)))
        For li = 1 To colWidth
            outCol = outCol + 1
            result(1, outCol) = colList(ci) & SuffixLetter(li)
        Next li
    Next ci

    ' 填充数据行
    Dim ri As Long, outRow As Long
    Dim rowDimKey As String
    Dim innerDict As Object
    Dim colKey As String
    Dim colItems As Object

    For ri = LBound(rowList) To UBound(rowList)
        outRow = ri - LBound(rowList) + 2
        result(outRow, 1) = rowList(ri)

        rowDimKey = rowList(ri)

        If Not groups.Exists(rowDimKey) Then
            FillRowPadding result, outRow, nColsOut, keepBlank
            GoTo NextOutRow
        End If

        Set innerDict = groups(rowDimKey)
        outCol = 1
        For ci = LBound(colList) To UBound(colList)
            colKey = colList(ci)
            colWidth = CLng(colMaxItems(colKey))

            If innerDict.Exists(colKey) Then
                Set colItems = innerDict(colKey)
            Else
                Set colItems = Nothing
            End If
            FillColumnCells result, outRow, outCol, colItems, colWidth, keepBlank
        Next ci

NextOutRow:
    Next ri
End Sub

'=============================================================================
' 辅助: 填充一列的数据单元格
'=============================================================================
Private Sub FillColumnCells(ByRef result As Variant, ByVal outRow As Long, _
                             ByRef outCol As Long, ByRef items As Object, _
                             ByVal colWidth As Long, ByVal keepBlank As Boolean)
    Dim li As Long
    For li = 1 To colWidth
        outCol = outCol + 1
        If Not items Is Nothing Then
            If li <= items.Count Then
                result(outRow, outCol) = items(li)
            ElseIf keepBlank Then
                result(outRow, outCol) = ""
            Else
                result(outRow, outCol) = Empty
            End If
        ElseIf keepBlank Then
            result(outRow, outCol) = ""
        Else
            result(outRow, outCol) = Empty
        End If
    Next li
End Sub

'=============================================================================
' 辅助: 填充整行的空白数据
'=============================================================================
Private Sub FillRowPadding(ByRef result As Variant, ByVal outRow As Long, _
                            ByVal nColsOut As Long, ByVal keepBlank As Boolean)
    Dim c As Long
    For c = 2 To nColsOut
        If keepBlank Then result(outRow, c) = "" Else result(outRow, c) = Empty
    Next c
End Sub

'=============================================================================
' 辅助: 将 Dictionary 键收集到字符串数组
'=============================================================================
Private Sub CollectKeys(ByRef dict As Object, ByRef keyList() As String)
    Dim keys As Variant
    keys = dict.Keys

    Dim lb As Long
    lb = LBound(keys)

    ReDim keyList(lb To UBound(keys))
    Dim i As Long
    For i = lb To UBound(keys)
        keyList(i) = CStr(keys(i))
    Next i
End Sub

'=============================================================================
' 辅助: 字符串数组升序排序
'=============================================================================
Private Sub SortStringArray(ByRef arr() As String)
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    If ub > lb Then QuickSortStrings arr, lb, ub
End Sub

Private Sub QuickSortStrings(ByRef arr() As String, ByVal low As Long, ByVal high As Long, _
                              Optional ByVal depth As Long = 0)
    If depth > MAX_SORT_DEPTH Then
        Err.Raise ERR_SORT_OVERFLOW, "QuickSortStrings", _
            "排序递归深度超过限制 (" & MAX_SORT_DEPTH & ")，数据量过大或数据分布异常。"
    End If
    ' 对小分区使用插入排序（n ≤ 16）
    If high - low <= 16 Then
        InsertionSortStrings arr, low, high
        Exit Sub
    End If

    Dim i As Long, j As Long
    Dim pivot As String, tmp As String

    ' Median-of-three pivot (SKILL.md §5.1) — avoids O(n²) on sorted/patterned data
    Dim mid As Long: mid = (low + high) \ 2
    If StrComp(arr(low), arr(mid), vbTextCompare) > 0 Then
        tmp = arr(low): arr(low) = arr(mid): arr(mid) = tmp
    End If
    If StrComp(arr(low), arr(high), vbTextCompare) > 0 Then
        tmp = arr(low): arr(low) = arr(high): arr(high) = tmp
    End If
    If StrComp(arr(mid), arr(high), vbTextCompare) > 0 Then
        tmp = arr(mid): arr(mid) = arr(high): arr(high) = tmp
    End If
    i = low: j = high
    pivot = arr(mid)
    Do While i <= j
        Do While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    ' 尾递归优化: 先递归较小的分区, 限制栈深度为 O(log n)
    If (j - low) < (high - i) Then
        If low < j Then QuickSortStrings arr, low, j, depth + 1
        If i < high Then QuickSortStrings arr, i, high, depth + 1
    Else
        If i < high Then QuickSortStrings arr, i, high, depth + 1
        If low < j Then QuickSortStrings arr, low, j, depth + 1
    End If
End Sub

' 小分区插入排序（n ≤ 16），避免 QuickSort 在小数组上的递归开销
Private Sub InsertionSortStrings(ByRef arr() As String, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long
    Dim key As String
    For i = low + 1 To high
        key = arr(i)
        j = i - 1
        Do While j >= low
            If StrComp(arr(j), key, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = key
    Next i
End Sub

'=============================================================================
' 辅助: 生成字母后缀 A, B, C, ..., Z, AA, AB, ...
'=============================================================================
Private Function SuffixLetter(ByVal n As Long) As String
    If n <= 0 Then
        SuffixLetter = ""
        Exit Function
    End If
    If n > MAX_EXCEL_COLS Then
        Err.Raise ERR_SUFFIX_RANGE, "SuffixLetter", _
            "后缀序号 (" & n & ") 超过 Excel 最大列数 " & MAX_EXCEL_COLS & "。"
    End If

    Dim result As String
    Dim remaining As Long
    Dim idx As Long
    remaining = n

    Do While remaining > 0
        idx = ((remaining - 1) Mod 26) + 1
        result = Mid$(ALPHABET, idx, 1) & result
        remaining = (remaining - 1) \ 26
    Loop

    SuffixLetter = result
End Function

'=============================================================================
' 公共辅助: 返回列名→列索引的字典
'=============================================================================

'=============================================================================

'=============================================================================
' 公共过程: 将透视结果直接输出到工作表 (自动适应输出范围)
'
' 用法:
'   RawConversionToRange Selection, "科目", "月份", "金额", Range("F1")
'   RawConversionToRange Range("A1:C100"), "姓名", "部门", "工资", Range("E1"), , True
'
' 参数:
'   srcRange    - 源数据区域 (含标题行)
'   valueColRef - 值填充列的标题文本 / 单元格引用 (其值填入交叉表单元格)
'   colDimRef   - 列维度列的标题文本 / 单元格引用 (其唯一值展开为列标题)
'   rowDimRef   - 行维度列的标题文本 / 单元格引用 (其唯一值作为输出行标签)
'   destCell    - 输出起始单元格 (结果将从此单元格开始写入)
'   keepBlank   - 是否保留空白值 (默认 True)
'   sortLabels  - 是否字母升序排序 (默认 False)
'=============================================================================
Public Sub RawConversionToRange( _
    ByVal srcRange As Range, _
    ByVal valueColRef As Variant, _
    ByVal colDimRef As Variant, _
    ByVal rowDimRef As Variant, _
    ByVal destCell As Range, _
    Optional ByVal keepBlank As Boolean = True, _
    Optional ByVal sortLabels As Boolean = False _
)
    On Error GoTo ErrHandler

    Dim result As Variant
    result = RawConversion(srcRange, valueColRef, colDimRef, rowDimRef, keepBlank, sortLabels)

    If IsError(result) Then
        destCell.Value = result
        Exit Sub
    End If

    ' 自动根据数组大小调整输出范围
    Dim nRows As Long, nCols As Long
    nRows = UBound(result, 1)
    nCols = UBound(result, 2)
    destCell.Resize(nRows, nCols).Value = result
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, ERR_SOURCE, Err.Description
End Sub

'=============================================================================
' Unpivot — 逆透视 (宽表 → 长表)
'
' 参数:
'   rng         - 源区域 (含标题行)
'   valueCols   - 需要逆透视的列 (值列)
'                 Variant: 可以是列索引数组如 Array(2,3,4)
'                 或第一个值列索引 (单值, 则后跟 countCols)
'   nameCol     - 属性名列标题文本 (如 "Month")
'   valCol      - 属性值列标题文本 (如 "Amount")
'   destCell    - 输出起始单元格
'   idColIndices - 保持不变的 ID 列 (可选, 默认第1列以外的都算值列)
'=============================================================================
Public Sub Unpivot( _
    ByRef rng As Range, _
    ByVal valueCols As Variant, _
    ByVal nameCol As String, _
    ByVal valCol As String, _
    ByRef destCell As Range, _
    Optional ByVal idColIndices As Variant)

    On Error GoTo ErrHandler

    If rng Is Nothing Then
        Err.Raise ERR_UNPIVOT_SRC, "Unpivot", "源数据区域不能为空。"
    End If
    If destCell Is Nothing Then
        Err.Raise ERR_UNPIVOT_DEST, "Unpivot", "输出起始单元格不能为空。"
    End If
    If rng.Rows.Count < 2 Then
        Err.Raise ERR_UNPIVOT_SRC, "Unpivot", "源数据区域至少需要2行 (标题 + 数据)。"
    End If

    Dim data As Variant
    If rng.Areas.Count > 1 Then Err.Raise ERR_MULTI_AREA, "Unpivot", "不支持多重选择区域。"
    data = rng.Value
    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1): nCols = UBound(data, 2)
    Dim i As Long, j As Long

    ' 解析值列和 ID 列
    Dim valColArr() As Long
    Dim idColArr() As Long
    Dim nValCols As Long, nIdCols As Long

    If IsArray(valueCols) Then
        nValCols = UBound(valueCols) - LBound(valueCols) + 1
        ReDim valColArr(1 To nValCols)
        For i = 1 To nValCols
            valColArr(i) = CLng(valueCols(LBound(valueCols) + i - 1))
        Next i
    Else
        nValCols = 1
        ReDim valColArr(1 To 1)
        valColArr(1) = CLng(valueCols)
    End If

    ' 验证值列索引
    For i = 1 To nValCols
        If valColArr(i) < 1 Or valColArr(i) > nCols Then
            Err.Raise ERR_OUT_OF_BOUNDS, "Unpivot", "值列索引 " & valColArr(i) & " 超出范围。"
        End If
    Next i

    If IsMissing(idColIndices) Then
        ReDim idColArr(1 To nCols)
        Dim isVal As Boolean
        Dim k As Long
        nIdCols = 0
        For j = 1 To nCols
            isVal = False
            For k = 1 To nValCols
                If j = valColArr(k) Then isVal = True: Exit For
            Next k
            If Not isVal Then
                nIdCols = nIdCols + 1
                idColArr(nIdCols) = j
            End If
        Next j
    Else
        If IsArray(idColIndices) Then
            nIdCols = UBound(idColIndices) - LBound(idColIndices) + 1
            ReDim idColArr(1 To nIdCols)
            For i = 1 To nIdCols
                idColArr(i) = CLng(idColIndices(LBound(idColIndices) + i - 1))
            Next i
        Else
            nIdCols = 1
            ReDim idColArr(1 To 1)
            idColArr(1) = CLng(idColIndices)
        End If
    End If

    ' 验证 ID 列索引
    For i = 1 To nIdCols
        If idColArr(i) < 1 Or idColArr(i) > nCols Then
            Err.Raise ERR_OUT_OF_BOUNDS, "Unpivot", "ID 列索引 " & idColArr(i) & " 超出范围。"
        End If
    Next i

    If nIdCols = 0 Then
        Err.Raise ERR_PIVOT_NO_ID, "Unpivot", "未找到 ID 列，所有列都被标记为值列。"
    End If
    If nValCols = 0 Then
        Err.Raise ERR_PIVOT_NO_VAL, "Unpivot", "未指定值列。"
    End If

    Dim outRows As Long
    outRows = 1 + (nRows - 1) * nValCols
    If outRows > MAX_EXCEL_ROWS Then
        Err.Raise ERR_PIVOT_OVERFLOW, "Unpivot", _
            "输出行数 (" & outRows & ") 超过 Excel 最大行数 " & MAX_EXCEL_ROWS & "。请减少值列或输入行。"
    End If
    Dim outCols As Long
    outCols = nIdCols + 2
    If outCols > MAX_EXCEL_COLS Then
        Err.Raise ERR_UNPIVOT_COLS, "Unpivot", _
            "输出列数 (" & outCols & ") 超过 Excel 最大列数 " & MAX_EXCEL_COLS & "。"
    End If

    Dim outArr() As Variant
    ReDim outArr(1 To outRows, 1 To outCols)

    Dim c As Long
    c = 1
    For j = 1 To nIdCols
        outArr(1, c) = data(1, idColArr(j))
        c = c + 1
    Next j
    outArr(1, c) = nameCol: c = c + 1
    outArr(1, c) = valCol

    Dim outRow As Long: outRow = 2
    For i = 2 To nRows
        For k = 1 To nValCols
            c = 1
            For j = 1 To nIdCols
                outArr(outRow, c) = data(i, idColArr(j))
                c = c + 1
            Next j
            outArr(outRow, c) = data(1, valColArr(k))
            c = c + 1
            outArr(outRow, c) = data(i, valColArr(k))
            outRow = outRow + 1
        Next k
    Next i

    destCell.Resize(outRows, outCols).Value = outArr
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, ERR_SOURCE, Err.Description
End Sub

'=============================================================================
' GroupBy — 分组聚合
'
' 参数:
'   rng      - 源区域 (含标题行)
'   groupCol - 分组列索引 (1-based)
'   aggCol   - 聚合列索引
'   aggFunc  - "SUM", "COUNT", "AVG", "MIN", "MAX" (默认 "SUM")
'   destCell - 输出起始单元格
'   groupColName - 分组列标题 (可选, 默认取源区域标题)
'=============================================================================
Public Function GroupBy( _
    ByRef rng As Range, _
    ByVal groupCol As Long, _
    ByVal aggCol As Long, _
    Optional ByVal aggFunc As String = "SUM", _
    Optional ByVal destCell As Range = Nothing) As Variant

    Dim data As Variant, nRows As Long, nCols As Long
    Dim dict As Object, dictCnt As Object, k As Variant

    On Error GoTo ErrHandler
    If rng Is Nothing Then
        Err.Raise ERR_INVALID_INPUT, "GroupBy", "Range 参数不能为 Nothing。"
    End If
    If rng.Rows.Count < 2 Then
        Err.Raise ERR_INVALID_INPUT, "GroupBy", "区域至少需要 2 行 (含标题行)。"
    End If

    If rng.Areas.Count > 1 Then Err.Raise ERR_MULTI_AREA, "GroupBy", "不支持多重区域。"
    data = rng.Value
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)

    If groupCol < 1 Or groupCol > nCols Then
        Err.Raise ERR_OUT_OF_BOUNDS, "GroupBy", "分组列越界: " & groupCol
    End If
    If aggCol < 1 Or aggCol > nCols Then
        Err.Raise ERR_OUT_OF_BOUNDS, "GroupBy", "聚合列越界: " & aggCol
    End If
    Set dict = DP.Create()
    dict.CompareMode = vbTextCompare

    Dim funcUpper As String
    funcUpper = UCase$(aggFunc)

    If funcUpper = "AVG" Then
        Set dictCnt = DP.Create()
        dictCnt.CompareMode = vbTextCompare
    End If

    Dim i As Long
    Dim groupKey As String, val As Double, isNum As Boolean

    For i = 2 To nRows
        If IsError(data(i, groupCol)) Or IsNull(data(i, groupCol)) Then GoTo ContinueRow
        groupKey = Trim(CStr(data(i, groupCol)))
        If Len(groupKey) = 0 Then GoTo ContinueRow

        If funcUpper <> "COUNT" Then
            If IsError(data(i, aggCol)) Then GoTo ContinueRow
            If VarType(data(i, aggCol)) = vbBoolean Or Not IsNumeric(data(i, aggCol)) Then GoTo ContinueRow
            val = CDbl(data(i, aggCol))
            isNum = True
        Else
            isNum = False
        End If

        Select Case funcUpper
            Case "COUNT"
                If dict.Exists(groupKey) Then
                    dict(groupKey) = CLng(dict(groupKey)) + 1
                Else
                    dict.Add groupKey, 1
                End If
            Case "SUM"
                If dict.Exists(groupKey) Then
                    dict(groupKey) = CDbl(dict(groupKey)) + val
                Else
                    dict.Add groupKey, val
                End If
            Case "AVG"
                If dict.Exists(groupKey) Then
                    dict(groupKey) = CDbl(dict(groupKey)) + val
                    dictCnt(groupKey) = CLng(dictCnt(groupKey)) + 1
                Else
                    dict.Add groupKey, val
                    dictCnt.Add groupKey, 1
                End If
            Case "MIN"
                If dict.Exists(groupKey) Then
                    If val < CDbl(dict(groupKey)) Then dict(groupKey) = val
                Else
                    dict.Add groupKey, val
                End If
            Case "MAX"
                If dict.Exists(groupKey) Then
                    If val > CDbl(dict(groupKey)) Then dict(groupKey) = val
                Else
                    dict.Add groupKey, val
                End If
            Case Else
                Err.Raise ERR_INVALID_INPUT, "GroupBy", "不支持的聚合函数: " & aggFunc
                Exit Function
        End Select
ContinueRow:
    Next i

    ' AVG 后处理
    If funcUpper = "AVG" Then
        For Each k In dict.Keys
            If dictCnt.Exists(k) Then
                If CLng(dictCnt(k)) > 0 Then
                    dict(k) = CDbl(dict(k)) / CLng(dictCnt(k))
                End If
            End If
        Next k
    End If

    If dict.Count = 0 Then
        Err.Raise ERR_EMPTY_DIM, "GroupBy", "分组结果为空 — 无有效聚合数据。"
    End If

    Dim outArr() As Variant
    ReDim outArr(1 To dict.Count + 1, 1 To 2)
    outArr(1, 1) = CStr(data(1, groupCol))
    outArr(1, 2) = funcUpper & "(" & CStr(data(1, aggCol)) & ")"

    Dim keys As Variant: keys = dict.Keys
    For i = 0 To dict.Count - 1
        outArr(i + 2, 1) = keys(i)
        outArr(i + 2, 2) = dict(keys(i))
    Next i

    If Not destCell Is Nothing Then
        destCell.Resize(dict.Count + 1, 2).Value = outArr
    End If

    GroupBy = outArr
    Set dict = Nothing
    Set dictCnt = Nothing
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "GroupBy", Err.Description
End Function

'=============================================================================
' SplitColumnToRows — 按分隔符拆分某列到多行
'
' 参数:
'   rng       - 源区域 (含标题行)
'   colIdx    - 要拆分的列索引 (1-based)
'   delimiter - 分隔符
'   destCell  - 输出起始单元格
'=============================================================================
Public Sub SplitColumnToRows( _
    ByRef rng As Range, _
    ByVal colIdx As Long, _
    ByVal delimiter As String, _
    ByRef destCell As Range)

    On Error GoTo ErrHandler

    If rng Is Nothing Then
        Err.Raise ERR_SPLIT_SRC, "SplitColumnToRows", "源数据区域不能为空。"
    End If
    If destCell Is Nothing Then
        Err.Raise ERR_SPLIT_DEST, "SplitColumnToRows", "输出起始单元格不能为空。"
    End If
    If Len(delimiter) = 0 Then
        Err.Raise ERR_SPLIT_DELIM, "SplitColumnToRows", "分隔符不能为空。"
    End If
    If rng.Rows.Count < 2 Then
        Err.Raise ERR_SPLIT_SRC, "SplitColumnToRows", "源数据区域至少需要2行 (标题 + 数据)。"
    End If

    Dim data As Variant
    If rng.Areas.Count > 1 Then Err.Raise ERR_MULTI_AREA, "SplitColumnToRows", "不支持多重选择区域。"
    data = rng.Value
    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1): nCols = UBound(data, 2)
    If colIdx < 1 Or colIdx > nCols Then
        Err.Raise ERR_OUT_OF_BOUNDS, "SplitColumnToRows", "列索引 " & colIdx & " 超出范围。"
    End If
    Dim i As Long, j As Long
    Dim parts() As String
    Dim pCnt As Long
    Dim p As Long, lb As Long

    ' 第一遍: 计算输出行数并缓存拆分结果
    Dim outRows As Long
    outRows = 1
    Dim cachedParts As Collection
    Set cachedParts = New Collection

    For i = 2 To nRows
        If IsError(data(i, colIdx)) Or IsNull(data(i, colIdx)) Then
            ReDim parts(0 To 0)
            parts(0) = ""
        Else
            parts = Split(CStr(data(i, colIdx)), delimiter)
        End If
        pCnt = UBound(parts) - LBound(parts) + 1
        If pCnt = 0 Then pCnt = 1
        outRows = outRows + pCnt
        cachedParts.Add parts
    Next i

    Dim outArr() As Variant
    ReDim outArr(1 To outRows, 1 To nCols)

    For j = 1 To nCols
        outArr(1, j) = data(1, j)
    Next j

    ' 数据行 — 复用缓存的拆分结果
    Dim outRow As Long: outRow = 2
    For i = 2 To nRows
        parts = cachedParts(i - 1)
        pCnt = UBound(parts) - LBound(parts) + 1

        If pCnt = 0 Then
            For j = 1 To nCols
                outArr(outRow, j) = data(i, j)
            Next j
            outRow = outRow + 1
        Else
            lb = LBound(parts)
            For p = 0 To pCnt - 1
                For j = 1 To nCols
                    If j = colIdx Then
                        outArr(outRow, j) = Trim(parts(lb + p))
                    Else
                        outArr(outRow, j) = data(i, j)
                    End If
                Next j
                outRow = outRow + 1
            Next p
        End If
    Next i

    destCell.Resize(outRows, nCols).Value = outArr
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, ERR_SOURCE, Err.Description
End Sub

'=============================================================================
' MergeColumns — 合并多列为一列
'
' 参数:
'   rng         - 源区域 (含标题行)
'   colIndices  - 要合并的列索引 (Variant 数组或单数)
'   delimiter   - 分隔符
'   newColName  - 新列标题
'   destCell    - 输出起始单元格
'   dropOrigCols - 是否删除原始列 (默认 False)
'=============================================================================
Public Sub MergeColumns( _
    ByRef rng As Range, _
    ByVal colIndices As Variant, _
    ByVal delimiter As String, _
    ByVal newColName As String, _
    ByRef destCell As Range, _
    Optional ByVal dropOrigCols As Boolean = False)

    On Error GoTo ErrHandler

    If rng Is Nothing Then
        Err.Raise ERR_MERGE_SRC, "MergeColumns", "源数据区域不能为空。"
    End If
    If destCell Is Nothing Then
        Err.Raise ERR_MERGE_DEST, "MergeColumns", "输出起始单元格不能为空。"
    End If
    If rng.Rows.Count < 2 Then
        Err.Raise ERR_MERGE_SRC, "MergeColumns", "源数据区域至少需要2行 (标题 + 数据)。"
    End If

    Dim data As Variant
    If rng.Areas.Count > 1 Then Err.Raise ERR_MULTI_AREA, "MergeColumns", "不支持多重选择区域。": Exit Sub
    data = rng.Value
    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1): nCols = UBound(data, 2)
    Dim i As Long, j As Long, c As Long
    Dim outArr() As Variant
    Dim mergeParts() As String
    Dim mpIdx As Long
    Dim mergeDict As Object
    Dim cellTmp As Variant
    Dim valStr As String
    Dim mi As Long

    ' 解析列索引
    Dim colArr() As Long
    Dim nMergeCols As Long
    If IsArray(colIndices) Then
        nMergeCols = UBound(colIndices) - LBound(colIndices) + 1
        ReDim colArr(1 To nMergeCols)
        For i = 1 To nMergeCols
            colArr(i) = CLng(colIndices(LBound(colIndices) + i - 1))
        Next i
    Else
        nMergeCols = 1
        ReDim colArr(1 To 1)
        colArr(1) = CLng(colIndices)
    End If

    ' 验证列索引范围
    Dim mergeIdx As Long
    For mergeIdx = 1 To nMergeCols
        If colArr(mergeIdx) < 1 Or colArr(mergeIdx) > nCols Then
            Err.Raise ERR_OUT_OF_BOUNDS, "MergeColumns", "列索引 " & colArr(mergeIdx) & " 超出范围。"
        End If
    Next mergeIdx

    Dim outCols As Long
    If dropOrigCols Then
        outCols = nCols - nMergeCols + 1
    Else
        outCols = nCols + 1
    End If

    ' Build merge-column lookup dict once (O(1) vs linear scan)
    Set mergeDict = DP.Create()
    For mergeIdx = 1 To nMergeCols
        mergeDict.Add colArr(mergeIdx), True
    Next mergeIdx

    ReDim outArr(1 To nRows, 1 To outCols)

    ReDim mergeParts(0 To nMergeCols)

    For i = 1 To nRows
        c = 1
        mpIdx = 0

        For j = 1 To nCols
            If mergeDict.Exists(j) Then
                cellTmp = data(i, j)
                If IsError(cellTmp) Or IsNull(cellTmp) Then cellTmp = ""
                valStr = Trim(CStr(cellTmp))
                If Len(valStr) > 0 Then
                    mergeParts(mpIdx) = valStr: mpIdx = mpIdx + 1
                End If
                If Not dropOrigCols Then
                    outArr(i, c) = data(i, j)
                    c = c + 1
                End If
            Else
                outArr(i, c) = data(i, j)
                c = c + 1
            End If
        Next j

        If i = 1 Then
            outArr(i, c) = newColName
        ElseIf mpIdx > 0 Then
            ' 固定大小缓冲 mergeParts(0 To nMergeCols) 全程不变 — 仅拼接前 mpIdx 个有效元素
            ' (修复原循环内 ReDim Preserve 收缩数组导致后续行写入越界 Error 9；SKILL.md §5.4: 用 Join 替代 & 循环)
            Dim tmpParts() As String
            ReDim tmpParts(0 To mpIdx - 1)
            For mi = 0 To mpIdx - 1
                tmpParts(mi) = mergeParts(mi)
            Next mi
            outArr(i, c) = Join(tmpParts, delimiter)
        Else
            outArr(i, c) = ""
        End If
    Next i

    destCell.Resize(nRows, outCols).Value = outArr
    Set mergeDict = Nothing
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, ERR_SOURCE, Err.Description
End Sub

'=============================================================================
' FilterTable — 按条件筛选表格 (包含标题行)
'
' 参数:
'   rng    - 源区域 (含标题行)
'   colIdx - 筛选列索引
'   op     - "=", "<", ">", "<=", ">=", "<>", "contains"
'   value  - 比较值
'   destCell - 输出起始单元格
'=============================================================================
Public Sub FilterTable( _
    ByRef rng As Range, _
    ByVal colIdx As Long, _
    ByVal op As String, _
    ByVal value As Variant, _
    ByRef destCell As Range)

    On Error GoTo ErrHandler

    If rng Is Nothing Then
        Err.Raise ERR_FILTER_SRC, "FilterTable", "源数据区域不能为空。"
    End If
    If destCell Is Nothing Then
        Err.Raise ERR_FILTER_DEST, "FilterTable", "输出起始单元格不能为空。"
    End If
    If rng.Rows.Count < 2 Then
        Err.Raise ERR_FILTER_SRC, "FilterTable", "源数据区域至少需要2行 (标题 + 数据)。"
    End If

    Dim result As Variant
    result = FilterRangeToArray(rng, colIdx, op, value)

    Dim nRows As Long, nCols As Long
    nRows = UBound(result, 1)
    nCols = UBound(result, 2)
    destCell.Resize(nRows, nCols).Value = result
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, ERR_SOURCE, Err.Description
End Sub

'=============================================================================
' 辅助: FilterRangeToArray — 按条件筛选表格 (返回二维数组)
'=============================================================================
Private Function FilterRangeToArray( _
    ByRef rng As Range, _
    ByVal colIdx As Long, _
    ByVal op As String, _
    ByVal value As Variant _
) As Variant

    Dim data As Variant
    If rng.Areas.Count > 1 Then FilterRangeToArray = CVErr(xlErrValue): Exit Function
    data = rng.Value
    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)

    If colIdx < 1 Or colIdx > nCols Then
        Err.Raise ERR_OUT_OF_BOUNDS, "FilterRangeToArray", "列索引 " & colIdx & " 超出范围。"
    End If

    ' 第一遍: 统计匹配行数
    Dim matchCnt As Long, i As Long
    matchCnt = 0
    Dim opUCase As String
    opUCase = UCase$(op)

    For i = 2 To nRows
        If IsError(data(i, colIdx)) Then GoTo ContinueCount
        If MatchCondition(data(i, colIdx), opUCase, value) Then
            matchCnt = matchCnt + 1
        End If
ContinueCount:
    Next i

    ' 构建输出数组 (标题行 + 匹配的数据行)
    Dim outArr() As Variant
    ReDim outArr(1 To matchCnt + 1, 1 To nCols)

    Dim j As Long
    For j = 1 To nCols
        outArr(1, j) = data(1, j)
    Next j

    Dim outRow As Long
    outRow = 2
    For i = 2 To nRows
        If Not IsError(data(i, colIdx)) Then
            If MatchCondition(data(i, colIdx), opUCase, value) Then
                For j = 1 To nCols
                    outArr(outRow, j) = data(i, j)
                Next j
                outRow = outRow + 1
            End If
        End If
    Next i

    FilterRangeToArray = outArr
End Function

'=============================================================================
' 辅助: MatchCondition — 委托给 VariantKit.FilterPasses (单一数据源)
'=============================================================================
Private Function MatchCondition( _
    ByVal cellVal As Variant, _
    ByVal opUCase As String, _
    ByVal cmpVal As Variant _
) As Boolean
    Static vk As VariantKit: If vk Is Nothing Then Set vk = New VariantKit
    MatchCondition = vk.FilterPasses(cellVal, cmpVal, opUCase)
    If Not MatchCondition Then
        ' FilterPasses 不识别时返回 False；检查是否为不支持的操作符
        Select Case UCase$(opUCase)
            Case "=", "<", ">", "<=", ">=", "<>", "CONTAINS"
                ' 识别但结果为 False — 正常
            Case Else
                Err.Raise ERR_FILTER_OP, "FilterTable", _
                    "不支持的操作符 """ & opUCase & """。支持: =, <, >, <=, >=, <>, CONTAINS"
        End Select
    End If
End Function

'=============================================================================
' TransposeTable — 表格行列转置 (保留标题位置)
'
' 将表格的第 1 行和第 1 列交换，相当于 Excel TRANSPOSE 但保持引用一致性
'=============================================================================
Public Sub TransposeTable(ByRef rng As Range, ByRef destCell As Range)
    On Error GoTo ErrHandler

    If rng Is Nothing Then
        Err.Raise ERR_TRANSTABLE_SRC, "TransposeTable", "源数据区域不能为空。"
    End If
    If destCell Is Nothing Then
        Err.Raise ERR_TRANSTABLE_DST, "TransposeTable", "输出起始单元格不能为空。"
    End If

    If rng.Count = 1 Then
        destCell.Value = rng.Value
        Exit Sub
    End If

    Dim data As Variant
    If rng.Areas.Count > 1 Then Err.Raise ERR_MULTI_AREA, "TransposeTable", "不支持多重选择区域。"
    data = rng.Value
    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1)
    nCols = UBound(data, 2)
    Dim i As Long, j As Long

    Dim outArr() As Variant
    ReDim outArr(1 To nCols, 1 To nRows)

    For i = 1 To nRows
        For j = 1 To nCols
            outArr(j, i) = data(i, j)
        Next j
    Next i

    destCell.Resize(nCols, nRows).Value = outArr
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, ERR_SOURCE, Err.Description
End Sub
' VLookupArray — 数组版 VLOOKUP (接受 Range 或二维数组)
'
' 参数:
'   dataArray   - 数据区域 (Range) 或包含标题行的二维数组
'   lookupValue - 要查找的值
'   lookupCol   - 查找列索引 (1-based)
'   returnCol   - 返回列索引 (1-based, 默认 2)
'
' 返回: 匹配行的 returnCol 列值。未找到/参数越界/无效输入时返回 CVErr。
' 查找为大小写不敏感 (vbTextCompare)。跳过 Null/Error 值。
'=============================================================================
'=============================================================================
Public Function VLookupArray( _
    ByVal dataArray As Variant, _
    ByVal lookupValue As Variant, _
    ByVal lookupCol As Long, _
    Optional ByVal returnCol As Long = 2) As Variant

    ' 从 Excel 传入 Range 时自动提取 .Value
    ' 守卫: TypeOf 仅适用于 Object，非对象输入需先检查 IsObject
    On Error GoTo ErrHandler

    Dim arrData As Variant
    Dim nRows As Long, i As Long, nCols As Long
    Dim lb As Long, cb As Long
    If IsObject(dataArray) Then
        If TypeOf dataArray Is Range Then
            arrData = dataArray.Value
        Else
            arrData = dataArray
        End If
    Else
        arrData = dataArray
    End If

    If IsEmpty(arrData) Or Not IsArray(arrData) Then
        Err.Raise ERR_INVALID_INPUT, "VLookupArray", "需要数组输入。"
    End If

    lb = LBound(arrData, 1)
    nRows = UBound(arrData, 1) - lb + 1
    cb = LBound(arrData, 2)
    nCols = UBound(arrData, 2) - cb + 1

    If lookupCol < 1 Or lookupCol > nCols Then
        Err.Raise ERR_OUT_OF_BOUNDS, "VLookupArray", "查找列越界: " & lookupCol
    End If
    If returnCol < 1 Or returnCol > nCols Then
        Err.Raise ERR_OUT_OF_BOUNDS, "VLookupArray", "返回列越界: " & returnCol
    End If

    For i = 2 To nRows
        If IsError(arrData(lb + i - 1, cb + lookupCol - 1)) Then GoTo ContinueVLookup
        If IsNull(arrData(lb + i - 1, cb + lookupCol - 1)) Or IsNull(lookupValue) Then GoTo ContinueVLookup
        If StrComp(CStr(arrData(lb + i - 1, cb + lookupCol - 1)), _
                   CStr(lookupValue), vbTextCompare) = 0 Then
            VLookupArray = arrData(lb + i - 1, cb + returnCol - 1)
            Exit Function
        End If
ContinueVLookup:
    Next i
    VLookupArray = CVErr(xlErrValue)
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "VLookupArray", Err.Description
End Function

'=============================================================================
' CrossJoin — 交叉连接两个表格 (笛卡尔积)
'
' 参数:
'   rng1 — 第一个区域 (含标题行)
'   rng2 — 第二个区域 (含标题行)
'
' 返回: 合并后的二维数组 (标题行 = rng1 标题 + rng2 标题)
' 输出行数 = 1 + (rng1数据行数) × (rng2数据行数)
' 超出 Excel 行数上限 (1,048,576) 或 VBA 整数范围时触发 Err.Raise
'=============================================================================
Public Function CrossJoin(ByRef rng1 As Range, ByRef rng2 As Range) As Variant
    On Error GoTo ErrHandler
    If rng1 Is Nothing Or rng2 Is Nothing Then
        CrossJoin = CVErr(xlErrValue): Exit Function
    End If
    If rng1.Rows.Count < 2 Or rng2.Rows.Count < 2 Then
        CrossJoin = CVErr(xlErrValue): Exit Function
    End If

    Dim data1 As Variant, data2 As Variant
    If rng1.Areas.Count > 1 Or rng2.Areas.Count > 1 Then CrossJoin = CVErr(xlErrValue): Exit Function
    data1 = rng1.Value: data2 = rng2.Value

    Dim lb1 As Long: lb1 = LBound(data1, 1)
    Dim lb2 As Long: lb2 = LBound(data2, 1)
    Dim nRows1 As Long: nRows1 = UBound(data1, 1) - lb1 + 1
    Dim cb1 As Long: cb1 = LBound(data1, 2)
    Dim cb2 As Long: cb2 = LBound(data2, 2)
    Dim nCols1 As Long: nCols1 = UBound(data1, 2) - cb1 + 1
    Dim nRows2 As Long: nRows2 = UBound(data2, 1) - lb2 + 1
    Dim nCols2 As Long: nCols2 = UBound(data2, 2) - cb2 + 1

    Dim dataRows1 As Long: dataRows1 = nRows1 - 1
    Dim dataRows2 As Long: dataRows2 = nRows2 - 1

    Dim outRows As Long
    Dim outCols As Long: outCols = nCols1 + nCols2

    ' 用 Double 检测溢出，避免 Long * Long 溢出后变负数绕过检查
    If CDbl(dataRows1) * CDbl(dataRows2) > CDbl(&H7FFFFFFF) Then
        CrossJoin = CVErr(xlErrValue): Exit Function
    End If
    outRows = 1 + dataRows1 * dataRows2

    If outRows > MAX_EXCEL_ROWS Then
        CrossJoin = CVErr(xlErrValue): Exit Function
    End If
    If outCols > MAX_EXCEL_COLS Then
        CrossJoin = CVErr(xlErrValue): Exit Function
    End If

    Dim outArr() As Variant
    ReDim outArr(1 To outRows, 1 To outCols)

    Dim c As Long, i As Long, j As Long, k As Long
    c = 1
    For j = 1 To nCols1: outArr(1, c) = data1(lb1, cb1 + j - 1): c = c + 1: Next j
    For j = 1 To nCols2: outArr(1, c) = data2(lb2, cb2 + j - 1): c = c + 1: Next j

    Dim outRow As Long: outRow = 2
    For i = 2 To nRows1
        For k = 2 To nRows2
            c = 1
            For j = 1 To nCols1
                outArr(outRow, c) = data1(lb1 + i - 1, cb1 + j - 1)
                c = c + 1
            Next j
            For j = 1 To nCols2
                outArr(outRow, c) = data2(lb2 + k - 1, cb2 + j - 1)
                c = c + 1
            Next j
            outRow = outRow + 1
        Next k
    Next i

    CrossJoin = outArr
    Exit Function

ErrHandler:
    Err.Raise Err.Number, "CrossJoin", Err.Description
End Function

'=============================================================================
' 工作表函数 (UDF_PIVOT_*) — 遵循 UDF_<模块简称>_<函数名> 命名规范
'=============================================================================

Public Function UDF_PIVOT_CONVERT(ByVal srcRange As Variant, ByVal valueColRef As Variant, _
    ByVal colDimRef As Variant, ByVal rowDimRef As Variant, _
    Optional ByVal keepBlank As Variant = True, Optional ByVal sortLabels As Variant = False) As Variant
    On Error GoTo EH
    If Not TypeOf srcRange Is Range Then UDF_PIVOT_CONVERT = CVErr(xlErrValue): Exit Function
    Dim pivRng As Range: Set pivRng = srcRange
    UDF_PIVOT_CONVERT = RawConversion(pivRng, valueColRef, colDimRef, rowDimRef, keepBlank, sortLabels)
    Exit Function
EH: UDF_PIVOT_CONVERT = CVErr(xlErrValue)
End Function

Public Function UDF_PIVOT_GROUPBY(ByVal rng As Variant, ByVal groupCol As Variant, _
    ByVal aggCol As Variant, Optional ByVal aggFunc As Variant = "SUM", _
    Optional ByVal destCell As Variant = Nothing) As Variant
    On Error GoTo EH
    If Not TypeOf rng Is Range Then UDF_PIVOT_GROUPBY = CVErr(xlErrValue): Exit Function
    Dim gbRng As Range: Set gbRng = rng
    ' UDF 上下文禁止写入工作表 (Excel 限制) — destCell 参数保留仅用于签名兼容性；
    ' 始终传 Nothing 给 GroupBy (若传入 destCell 则 UDF 会 Error 1004)。
    ' 聚合结果通过动态数组溢出返回。
    UDF_PIVOT_GROUPBY = GroupBy(gbRng, groupCol, aggCol, aggFunc, Nothing)
    Exit Function
EH: UDF_PIVOT_GROUPBY = CVErr(xlErrValue)
End Function

Public Function UDF_PIVOT_VLOOKUP(ByVal dataArray As Variant, ByVal lookupValue As Variant, _
    ByVal lookupCol As Variant, Optional ByVal returnCol As Variant = 2) As Variant
    On Error GoTo EH: UDF_PIVOT_VLOOKUP = VLookupArray(dataArray, lookupValue, lookupCol, returnCol): Exit Function
EH: UDF_PIVOT_VLOOKUP = CVErr(xlErrValue)
End Function

Public Function UDF_PIVOT_CROSSJOIN(ByVal rng1 As Variant, ByVal rng2 As Variant) As Variant
    On Error GoTo EH
    If Not TypeOf rng1 Is Range Then UDF_PIVOT_CROSSJOIN = CVErr(xlErrValue): Exit Function
    If Not TypeOf rng2 Is Range Then UDF_PIVOT_CROSSJOIN = CVErr(xlErrValue): Exit Function
    Dim cjRng1 As Range: Set cjRng1 = rng1
    Dim cjRng2 As Range: Set cjRng2 = rng2
    UDF_PIVOT_CROSSJOIN = CrossJoin(cjRng1, cjRng2)
    Exit Function
EH: UDF_PIVOT_CROSSJOIN = CVErr(xlErrValue)
End Function
