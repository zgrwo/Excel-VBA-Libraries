Option Explicit

'==============================================================================
' Module:       SqlUtils
' Purpose:      SQL queries on Excel ranges via ADODB
' Layer:        Data
' Dependencies: ADODB (Windows built-in)
' Public:       18 functions/subs
' Notes:        Requires ADODB. 64-bit Office needs Access Database Engine 2016.
'==============================================================================


'=====================================================================
' SqlUtils.bas — SQL 查询工具集
'
' 通过 ADODB (ACE/Jet OLEDB) 将 Excel 工作表当作数据库表查询。
' 支持当前工作簿、外部工作簿、命名区域作为数据源。
'
' ⚠️ 模块缓存 ADODB 连接以提升性能。请在 Workbook_BeforeClose
'    中调用 CloseSqlCache 释放连接，避免 Excel 进程残留。
'
' 工作表函数 (UDF_SQL_*):
'   UDF_SQL_QUERY        — 执行 SELECT 查询，返回结果数组
'   UDF_SQL_JOIN         — 两表 JOIN
'   UDF_SQL_GROUPBY      — 分组聚合
'   UDF_SQL_LIST_SHEETS  — 列出工作簿中的所有工作表
'   UDF_SQL_LIST_COLUMNS — 列出表的列名
'   UDF_SQL_LIST_TABLES  — 列出可用数据源
'
' 内部函数 (PascalCase，供 VBA 代码调用):
'   SqlExecute           — 执行任意 SQL，返回 2D Variant 数组
'   SqlQuery             — SELECT 查询（自动推断数据源）
'   SqlJoin              — 两表 JOIN
'   SqlGroupBy           — 分组聚合
'   SqlListSheets        — 列出所有工作表名
'   SqlListColumns       — 列出指定表的列名
'   SqlListTables        — 列出所有可用数据源
'   SqlRangeQuery        — 对 Range 直接查询（无需保存工作簿）
'
' 依赖: ADODB (Windows 内置), ACE.OLEDB.12.0 / Jet.OLEDB.4.0
'
' 限制:
'   - ACE.OLEDB 在 64 位 Office 中需安装 Access Database Engine
'   - Jet.OLEDB 仅支持 32 位且 .xls 格式
'   - 工作表名含特殊字符时自动加方括号转义
'
' 错误码:
'   vbObjectError+1201  ERR_NOT_AVAIL     — ADODB/ACE/Jet 不可用
'   vbObjectError+1202  ERR_INVALID_INPUT — 空 SQL、Nothing Range、未保存工作簿
'   vbObjectError+1203  ERR_SQL_ERROR     — SQL 执行错误 (保留)
'
' ACE/Jet 常见故障:
'   "未找到提供程序"        → 安装 Access Database Engine 或改用 32 位 Office
'   "文件无效"              → 文件路径含特殊字符，或 .xls 用 ACE 打开
'   "找不到对象"            → 表名/工作表名不存在，检查 $ 后缀
'   "字段定义语法错误"       → 列名含非法字符，检查 MakeSafeColumnName
'
' 资源释放: 调用 CloseSqlCache 或在 Workbook_BeforeClose 中释放连接
'
' ⚠️ SQL 注入风险:
'   ACE OLEDB Excel ISAM 驱动不支持参数化查询 (ADODB.Command.Parameters)。
'   所有查询通过字符串拼接构建，用户提供的值直接拼入 SQL 语句。
'   使用 SqlEscapeString() 转义单引号可防止基本的注入攻击:
'     =UDF_SQL_QUERY("SELECT * FROM [Sheet1$] WHERE Name = '" & SqlEscapeString(A1) & "'")
'   注意: ACE/Jet 不支持多语句、DDL、UNION (部分版本)，降低了攻击面。
'=====================================================================

Private Const ERR_NOT_AVAIL     As Long = vbObjectError + 1201
Private Const ERR_INVALID_INPUT As Long = vbObjectError + 1202
Private Const ERR_SQL_ERROR     As Long = vbObjectError + 1203

' 模块级连接缓存 (避免重复 Open 开销)
Private mCachedConn As Object
Private mCachedPath As String

Private Const ACE_CONN_PREFIX   As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
Private Const ACE16_CONN_PREFIX As String = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source="
Private Const ACE_EXT_PROP      As String = ";Extended Properties=""Excel 12.0;HDR=YES;IMEX=1"""
Private Const JET_CONN_PREFIX   As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Private Const JET_EXT_PROP      As String = ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""

'=============================================================================
' SqlEscapeString — 转义 SQL 字符串中的单引号，防止 SQL 注入
'
' 用法: 在拼接用户输入到 WHERE/VALUES 子句时使用:
'   sql = "... WHERE Name = '" & SqlEscapeString(userInput) & "'"
'
' @Function Public Function SqlEscapeString(ByVal value As Variant, Optional ByVal forLike As Boolean = False) As String
' @Description 将单引号转义为两个单引号（SQL 标准转义）。forLike=True 时额外转义 LIKE 通配符。
' @Args value: 需要转义的字符串。非字符串类型自动转换为字符串。
' @Args forLike: 是否按 ACE LIKE 语法转义通配符 (% _ [) — 默认 False
' @Returns 单引号被转义为 '' 的字符串；forLike=True 时同时转义 LIKE 特殊字符
'
' ⚠️ 适用路径 (R5-26): 方括号字符类转义仅对 ACE/Jet SQL 文本执行路径合法
'   (SqlExecute / SqlQuery / SqlJoin / SqlGroupBy)。SqlRangeQuery 的 WHERE
'   子句由 ADO Recordset.Filter 执行 — 其 LIKE 引擎把 [ / ] 当字面量, 且
'   不支持 IN / BETWEEN, 详见 SqlRangeQuery 说明。
'=============================================================================
Public Function SqlEscapeString(ByVal value As Variant, Optional ByVal forLike As Boolean = False) As String
    If IsNull(value) Or IsEmpty(value) Then
        SqlEscapeString = ""
        Exit Function
    End If
    ' Guard: CStr on Error variants (#N/A, #VALUE!) raises Error 13
    If VarType(value) = vbError Then
        SqlEscapeString = ""
        Exit Function
    End If
    SqlEscapeString = Replace(CStr(value), "'", "''")
    If forLike Then
        ' ACE/Jet LIKE 使用方括号字符类作为转义语法; 反斜杠不是转义符
        ' (实测: LIKE '100\%' 不匹配 '100%', LIKE 'a\[b' 直接报无效模式串).
        ' 顺序: 先转义 [ 自身, 再包住通配符; ] 单独出现即字面量, 无需处理.
        SqlEscapeString = Replace(SqlEscapeString, "[", "[[]")
        SqlEscapeString = Replace(SqlEscapeString, "%", "[%]")
        SqlEscapeString = Replace(SqlEscapeString, "_", "[_]")
    End If
End Function

'=============================================================================
' CloseSqlCache — 释放缓存的数据库连接
'=============================================================================
Public Sub CloseSqlCache()
    If Not mCachedConn Is Nothing Then
        On Error Resume Next
        mCachedConn.Close
        On Error GoTo 0
        Set mCachedConn = Nothing: mCachedPath = ""
    End If
End Sub

'=============================================================================
' SqlGetConnection — 获取 ADODB 连接 (Static 缓存，复用同一文件的连接)
'=============================================================================
Public Function SqlGetConnection( _
    Optional ByVal filePath As String = "", _
    Optional ByRef outOk As Boolean) As Object

    Dim conn As Object, connStr As String, srcPath As String
    outOk = False

    If Len(filePath) = 0 Then
        Dim defWb As Workbook
        Set defWb = ThisWorkbook
        ' .xlam 加载项自身无法被 ACE 打开 — 默认改用活动工作簿
        If defWb.IsAddIn Then
            ' Set 失败时对象变量保留原值 (R5-38), 先置 Nothing 使守卫可达
            Set defWb = Nothing
            On Error Resume Next
            Set defWb = ActiveWorkbook
            On Error GoTo 0
        End If
        If defWb Is Nothing Then
            Err.Raise ERR_INVALID_INPUT, "SqlUtils", _
                "无可用的活动工作簿。请显式传入 filePath，或先打开目标工作簿。"
        End If
        srcPath = defWb.FullName
        If Len(srcPath) = 0 Then
            Err.Raise ERR_INVALID_INPUT, "SqlUtils", _
                "工作簿尚未保存。请先保存工作簿，或使用 SqlRangeQuery 对 Range 查询。"
        End If
    Else
        srcPath = filePath
    End If

    ' 防止连接字符串注入 — 文件路径不得包含分号或双引号
    If InStr(srcPath, ";") > 0 Or InStr(srcPath, """") > 0 Then
        Err.Raise ERR_INVALID_INPUT, "SqlGetConnection", _
            "文件路径包含无效字符 — " & srcPath
    End If

    ' 复用模块级缓存的连接 (同一文件路径)
    If Not mCachedConn Is Nothing Then
        If StrComp(mCachedPath, srcPath, vbTextCompare) = 0 Then
            On Error Resume Next
            If mCachedConn.State = 1 Then  ' adStateOpen
                On Error GoTo 0
                outOk = True
                Set SqlGetConnection = mCachedConn
                Exit Function
            End If
            Err.Clear: On Error GoTo 0
        End If
        ' 路径不同或连接已失效 — 关闭旧的
        On Error Resume Next
        mCachedConn.Close
        On Error GoTo 0
        Set mCachedConn = Nothing: mCachedPath = ""
    End If

    ' 优先 ACE 12 → ACE 16 (仅 ADE 2016 环境) → Jet (.xls, 32 位)
    connStr = ACE_CONN_PREFIX & srcPath & ACE_EXT_PROP
    On Error Resume Next
    Set conn = CreateObject("ADODB.Connection")
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        Err.Raise ERR_NOT_AVAIL, "SqlUtils", "ADODB 不可用。"
    End If
    conn.Open connStr
    If Err.Number = 0 Then
        On Error GoTo 0
        outOk = True
        Set SqlGetConnection = conn
        Set mCachedConn = conn: mCachedPath = srcPath
        Exit Function
    End If

    ' 回退 ACE 16 (64 位环境通常仅注册 16.0)
    Dim aceErrDesc As String: aceErrDesc = Err.Description
    Err.Clear
    conn.Open ACE16_CONN_PREFIX & srcPath & ACE_EXT_PROP
    If Err.Number = 0 Then
        On Error GoTo 0
        outOk = True
        Set SqlGetConnection = conn
        Set mCachedConn = conn: mCachedPath = srcPath
        Exit Function
    End If

    ' 回退 Jet (仅 .xls)
    Dim ace16ErrDesc As String: ace16ErrDesc = Err.Description
    Err.Clear
    connStr = JET_CONN_PREFIX & srcPath & JET_EXT_PROP
    conn.Open connStr
    If Err.Number = 0 Then
        On Error GoTo 0
        outOk = True
        Set SqlGetConnection = conn
        Set mCachedConn = conn: mCachedPath = srcPath
        Exit Function
    End If

    Dim jetErrDesc As String: jetErrDesc = Err.Description
    Err.Clear: On Error GoTo 0
    If Not conn Is Nothing Then
        On Error Resume Next: conn.Close: On Error GoTo 0
        Set conn = Nothing
    End If
    Dim bitHint As String
#If Win64 Then
    bitHint = "当前为 64 位 Office — 请安装 64 位 Access Database Engine 2010/2016。"
#Else
    bitHint = "当前为 32 位 Office — 请安装 32 位 Access Database Engine 2010/2016。"
#End If
    Err.Raise ERR_NOT_AVAIL, "SqlUtils", _
        "无法建立数据库连接。" & vbLf & _
        "文件: " & srcPath & vbLf & _
        "ACE.12: " & aceErrDesc & vbLf & _
        "ACE.16: " & ace16ErrDesc & vbLf & _
        "Jet: " & jetErrDesc & vbLf & bitHint
End Function

'=============================================================================
' EscapeSheetName — 转义工作表名
'=============================================================================
Private Function EscapeSheetName(ByVal sheetName As String) As String
    If Left$(sheetName, 1) = "[" Then
        If Right$(sheetName, 1) <> "]" Then
            Err.Raise ERR_INVALID_INPUT, "EscapeSheetName", "工作表名方括号不配对: " & sheetName
        End If
        EscapeSheetName = sheetName: Exit Function
    End If
    ' Escape ] as ]] (SQL standard for bracket-delimited identifiers)
    sheetName = Replace(sheetName, "]", "]]")
    If Right$(sheetName, 1) = "$" Then
        EscapeSheetName = "[" & sheetName & "]"
    Else
        EscapeSheetName = "[" & sheetName & "$]"
    End If
End Function

'=============================================================================
' NormalizeTableName — 数据源名归一化 (OpenSchema 匹配用)
'
' 统一以下入参/返回值形式为不带包裹的 `Name$`:
'   Data Sheet           → Data Sheet$
'   'Data Sheet$'        → Data Sheet$   (ACE OpenSchema 对含空格名返回带单引号)
'   Data Sheet$          → Data Sheet$
'   [Data Sheet$]        → Data Sheet$
' 保持 $ 后缀语义; 空串归一化为 "$" (不会匹配任何实际表)。
'=============================================================================
Private Function NormalizeTableName(ByVal tableName As String) As String
    Dim tbl As String: tbl = Trim$(tableName)
    ' 剥离成对的方括号标识符
    If Len(tbl) >= 2 Then
        If Left$(tbl, 1) = "[" And Right$(tbl, 1) = "]" Then
            tbl = Mid$(tbl, 2, Len(tbl) - 2)
        End If
    End If
    ' 剥离成对的单引号 (R5-27: ACE OpenSchema 返回 'Data Sheet$')
    If Len(tbl) >= 2 Then
        If Left$(tbl, 1) = "'" And Right$(tbl, 1) = "'" Then
            tbl = Mid$(tbl, 2, Len(tbl) - 2)
        End If
    End If
    If Right$(tbl, 1) <> "$" Then tbl = tbl & "$"
    NormalizeTableName = tbl
End Function

'=============================================================================
' MakeSafeColumnName — 清理列名/标识符，移除 SQL 特殊字符
'
' 用法: 在动态构建 SQL 时清理用户提供的列名/表名:
'   sql = "SELECT [" & MakeSafeColumnName(userCol) & "] FROM [Sheet1$]"
'
' @Function Public Function MakeSafeColumnName(ByVal colName As String) As String
' @Description 移除方括号、$、单引号等可能导致 SQL 注入或语法错误的字符。
' @Args colName: 需要清理的标识符名称
' @Returns 清理后的安全标识符
'=============================================================================
Public Function MakeSafeColumnName(ByVal colName As String) As String
    colName = Replace(colName, "[", "")
    colName = Replace(colName, "]", "")
    colName = Replace(colName, "$", "_")
    colName = Replace(colName, "'", "")
    If Len(Trim(colName)) = 0 Then colName = "F"
    MakeSafeColumnName = colName
End Function

'=============================================================================
' SqlExecute — 执行 SQL 语句，返回 2D Variant 数组 (1-based, 含列名)
'=============================================================================
Public Function SqlExecute( _
    ByVal sql As String, _
    Optional ByVal filePath As String = "", _
    Optional ByVal includeHeader As Boolean = True, _
    Optional ByRef outOk As Boolean, _
    Optional ByRef outErrorMsg As String) As Variant()

    Dim conn As Object, rs As Object
    Dim result() As Variant
    Dim nRows As Long, nCols As Long, i As Long, j As Long
    Dim rawData As Variant, rowOffset As Long
    Dim savedErrNum As Long, savedErrDesc As String

    outOk = False
    If Len(Trim(sql)) = 0 Then
        Err.Raise ERR_INVALID_INPUT, "SqlUtils", "SQL 语句不能为空。"
    End If

    Set conn = SqlGetConnection(filePath, outOk)
    If Not outOk Then Exit Function

    On Error GoTo SqlError
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1, 1 ' adOpenForwardOnly, adLockReadOnly, adCmdText

    If rs.EOF Then
        ' 1x1 Empty array — Excel shows blank cell instead of #VALUE!
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = Empty
        SqlExecute = result: outOk = True
        GoTo Cleanup
    End If

    nCols = rs.Fields.Count
    rawData = rs.GetRows()                    ' → (0..nCols-1, 0..nRows-1)
    nRows = UBound(rawData, 2) + 1
    If includeHeader Then rowOffset = 1 Else rowOffset = 0
    ReDim result(1 To nRows + rowOffset, 1 To nCols)

    If includeHeader Then
        For j = 0 To nCols - 1
            result(1, j + 1) = rs.Fields(j).Name
        Next j
    End If
    For i = 0 To nRows - 1
        For j = 0 To nCols - 1
            If IsNull(rawData(j, i)) Then
                result(i + 1 + rowOffset, j + 1) = Empty
            Else
                result(i + 1 + rowOffset, j + 1) = rawData(j, i)
            End If
        Next j
    Next i

    SqlExecute = result: outOk = True
    GoTo Cleanup

SqlError:
    outOk = False
    savedErrNum = Err.Number
    savedErrDesc = Err.Description
    outErrorMsg = "[" & CStr(savedErrNum) & "] " & savedErrDesc
    Resume Cleanup

Cleanup:
    If Not rs Is Nothing Then
        On Error Resume Next: rs.Close: On Error GoTo 0
        Set rs = Nothing
    End If
    ' Do NOT close conn — it may be the cached connection from SqlGetConnection.
    ' Only release the local reference. CloseSqlCache handles the actual cleanup.
    Set conn = Nothing
End Function

'=============================================================================
' SqlQuery — SELECT 查询（分句传入，自动拼接）
'
' ⚠️ 安全警告: selectClause / whereClause / orderByClause 作为原始 SQL 片段直接
' 拼入查询语句。ACE OLEDB Excel ISAM 驱动不支持参数化查询 (ADODB.Command.Parameters)。
' 调用者必须自行确保这些子句中不包含未转义的用户输入。
' 对于字符串值, 使用 SqlEscapeString() 转义单引号; 对于 LIKE 模式, 使用 forLike:=True。
'=============================================================================
Public Function SqlQuery( _
    ByVal selectClause As String, _
    ByVal fromClause As String, _
    Optional ByVal whereClause As String = "", _
    Optional ByVal orderByClause As String = "", _
    Optional ByVal filePath As String = "", _
    Optional ByRef outOk As Boolean) As Variant()

    Dim sql As String
    ' Auto-escape plain table names (not already escaped with brackets)
    Dim safeFrom As String: safeFrom = fromClause
    If InStr(safeFrom, "[") = 0 Then
        safeFrom = EscapeSheetName(safeFrom)
    End If
    sql = "SELECT " & selectClause & " FROM " & safeFrom
    If Len(whereClause) > 0 Then sql = sql & " WHERE " & whereClause
    If Len(orderByClause) > 0 Then sql = sql & " ORDER BY " & orderByClause
    SqlQuery = SqlExecute(sql, filePath, True, outOk)
End Function

'=============================================================================
' SqlJoin — 两表 JOIN
'=============================================================================
Public Function SqlJoin( _
    ByVal table1 As String, _
    ByVal table2 As String, _
    ByVal joinOn As String, _
    Optional ByVal joinType As String = "INNER", _
    Optional ByVal selectCols As String = "*", _
    Optional ByVal filePath As String = "", _
    Optional ByRef outOk As Boolean) As Variant()

    Dim sql As String
    Select Case UCase$(joinType)
        Case "INNER", "LEFT", "RIGHT", "FULL": joinType = UCase$(joinType)
        Case Else: Err.Raise ERR_INVALID_INPUT, "SqlJoin", "不支持的 JOIN 类型: " & joinType & "。支持 INNER/LEFT/RIGHT/FULL。"
    End Select
    sql = "SELECT " & selectCols & " FROM " & EscapeSheetName(table1) & " AS t1 " & _
          joinType & " JOIN " & EscapeSheetName(table2) & " AS t2 ON " & joinOn
    SqlJoin = SqlExecute(sql, filePath, True, outOk)
End Function

'=============================================================================
' SqlGroupBy — 分组聚合
'=============================================================================
Public Function SqlGroupBy( _
    ByVal tableName As String, _
    ByVal groupCols As String, _
    ByVal aggExprs As String, _
    Optional ByVal whereClause As String = "", _
    Optional ByVal filePath As String = "", _
    Optional ByRef outOk As Boolean) As Variant()

    Dim sql As String
    sql = "SELECT " & groupCols & ", " & aggExprs & _
          " FROM " & EscapeSheetName(tableName)
    If Len(whereClause) > 0 Then sql = sql & " WHERE " & whereClause
    sql = sql & " GROUP BY " & groupCols
    SqlGroupBy = SqlExecute(sql, filePath, True, outOk)
End Function

'=============================================================================
' SqlListSheets — 列出工作簿中的所有工作表名 (OpenSchema)
'=============================================================================
Public Function SqlListSheets( _
    Optional ByVal filePath As String = "", _
    Optional ByRef outOk As Boolean) As Variant()

    Dim conn As Object, rs As Object
    Dim result() As Variant, i As Long

    outOk = False
    On Error GoTo ErrSchema
    Set conn = SqlGetConnection(filePath, outOk)
    If Not outOk Then Exit Function
    Set rs = conn.OpenSchema(20)  ' adSchemaTables
    rs.Filter = "TABLE_TYPE='TABLE'"  ' 排除 Excel 内部系统表 (如 _xlnm#_FilterDatabase)
    If rs.EOF Then
        ReDim result(1 To 1, 1 To 1): result(1, 1) = Empty
        SqlListSheets = result: outOk = True: GoTo Cleanup
    End If

    Dim rows As Variant: rows = rs.GetRows()
    Dim n As Long: n = UBound(rows, 2) + 1
    ' R5-28: 打开中的 .xlsm 经 OpenSchema 返回重复表/列, 按 TABLE_NAME 去重
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    For i = 0 To n - 1
        If Not IsNull(rows(2, i)) Then  ' TABLE_NAME = field index 2
            If Not seen.Exists(CStr(rows(2, i))) Then seen.Add CStr(rows(2, i)), True
        End If
    Next i
    Dim keys As Variant: keys = seen.Keys
    ReDim result(1 To seen.Count + 1, 1 To 1)
    result(1, 1) = "SheetName"
    For i = 0 To seen.Count - 1
        result(i + 2, 1) = keys(i)
    Next i
    SqlListSheets = result: outOk = True
    GoTo Cleanup

ErrSchema:
    outOk = False
Cleanup:
    If Not rs Is Nothing Then
        On Error Resume Next: rs.Close: On Error GoTo 0: Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        On Error Resume Next: On Error GoTo 0: Set conn = Nothing  ' cached conn — do not close
    End If
End Function

'=============================================================================
' SqlListColumns — 列出指定表的列名 (OpenSchema)
'=============================================================================
Public Function SqlListColumns( _
    ByVal tableName As String, _
    Optional ByVal filePath As String = "", _
    Optional ByRef outOk As Boolean) As Variant()

    Dim conn As Object, rs As Object
    Dim result() As Variant, i As Long, cnt As Long
    ' Normalize: 剥离 [] / ACE OpenSchema 单引号包裹 ('Name$'), 保持 $ 后缀 (R5-27)
    Dim tbl As String: tbl = NormalizeTableName(tableName)

    outOk = False
    On Error GoTo ErrSchema
    Set conn = SqlGetConnection(filePath, outOk)
    If Not outOk Then Exit Function

    Set rs = conn.OpenSchema(4)  ' adSchemaColumns
    If rs.EOF Then
        ReDim result(1 To 1, 1 To 1): result(1, 1) = Empty
        SqlListColumns = result: outOk = True: GoTo Cleanup
    End If

    ' ACE 对含空格表名返回 'Data Sheet$' (单引号是字段值的一部分), rs.Filter 的
    ' 字符串字面量无法同时兼容带/不带引号的 provider 行为 — 逐行归一化后比较 (R5-27)
    ' R5-28: 同一表在打开中的 .xlsm 下返回重复列, 按 ORDINAL_POSITION 去重
    cnt = 0: rs.MoveFirst
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        If Not IsNull(rs.Fields("TABLE_NAME").Value) Then
            If StrComp(NormalizeTableName(CStr(rs.Fields("TABLE_NAME").Value)), tbl, vbTextCompare) = 0 Then
                If Not IsNull(rs.Fields("ORDINAL_POSITION").Value) Then
                    If Not seen.Exists(CStr(rs.Fields("ORDINAL_POSITION").Value)) Then
                        seen.Add CStr(rs.Fields("ORDINAL_POSITION").Value), True
                        cnt = cnt + 1
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Loop
    If cnt = 0 Then
        ReDim result(1 To 1, 1 To 1): result(1, 1) = Empty
        SqlListColumns = result: outOk = True: GoTo Cleanup
    End If
    rs.MoveFirst

    Dim emitted As Object: Set emitted = CreateObject("Scripting.Dictionary")
    ReDim result(1 To cnt + 1, 1 To 2)
    result(1, 1) = "ColIndex": result(1, 2) = "ColName"
    i = 0
    Do While Not rs.EOF
        If Not IsNull(rs.Fields("TABLE_NAME").Value) Then
            If StrComp(NormalizeTableName(CStr(rs.Fields("TABLE_NAME").Value)), tbl, vbTextCompare) = 0 Then
                If Not IsNull(rs.Fields("ORDINAL_POSITION").Value) Then
                    If Not emitted.Exists(CStr(rs.Fields("ORDINAL_POSITION").Value)) Then
                        emitted.Add CStr(rs.Fields("ORDINAL_POSITION").Value), True
                        i = i + 1
                        result(i + 1, 1) = rs.Fields("ORDINAL_POSITION").Value
                        result(i + 1, 2) = rs.Fields("COLUMN_NAME").Value
                    End If
                End If
            End If
        End If
        rs.MoveNext
    Loop
    SqlListColumns = result: outOk = True
    GoTo Cleanup

ErrSchema:
    outOk = False
Cleanup:
    If Not rs Is Nothing Then
        On Error Resume Next: rs.Close: On Error GoTo 0: Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        On Error Resume Next: On Error GoTo 0: Set conn = Nothing  ' cached conn — do not close
    End If
End Function

'=============================================================================
' SqlListTables — 列出所有可用数据源 (OpenSchema)
'=============================================================================
Public Function SqlListTables( _
    Optional ByVal filePath As String = "", _
    Optional ByRef outOk As Boolean) As Variant()

    Dim conn As Object, rs As Object
    Dim result() As Variant, i As Long, cnt As Long

    outOk = False
    On Error GoTo ErrSchema
    Set conn = SqlGetConnection(filePath, outOk)
    If Not outOk Then Exit Function

    Set rs = conn.OpenSchema(20)  ' adSchemaTables
    If rs.EOF Then
        ReDim result(1 To 1, 1 To 1): result(1, 1) = Empty
        SqlListTables = result: outOk = True: GoTo Cleanup
    End If

    ' Count rows (R5-28: 按 TABLE_NAME 去重, 打开中的 .xlsm 会返回重复表)
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    cnt = 0: rs.MoveFirst
    Do While Not rs.EOF
        If Not IsNull(rs.Fields("TABLE_NAME").Value) Then
            If Not seen.Exists(CStr(rs.Fields("TABLE_NAME").Value)) Then
                seen.Add CStr(rs.Fields("TABLE_NAME").Value), True
                cnt = cnt + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.MoveFirst

    Dim emitted As Object: Set emitted = CreateObject("Scripting.Dictionary")
    ReDim result(1 To cnt + 1, 1 To 2)
    result(1, 1) = "TableName": result(1, 2) = "TableType"
    i = 0
    Do While Not rs.EOF
        If Not IsNull(rs.Fields("TABLE_NAME").Value) Then
            If Not emitted.Exists(CStr(rs.Fields("TABLE_NAME").Value)) Then
                emitted.Add CStr(rs.Fields("TABLE_NAME").Value), True
                i = i + 1
                result(i + 1, 1) = rs.Fields("TABLE_NAME").Value
                result(i + 1, 2) = rs.Fields("TABLE_TYPE").Value
            End If
        End If
        rs.MoveNext
    Loop
    SqlListTables = result: outOk = True
    GoTo Cleanup

ErrSchema:
    outOk = False
Cleanup:
    If Not rs Is Nothing Then
        On Error Resume Next: rs.Close: On Error GoTo 0: Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        On Error Resume Next: On Error GoTo 0: Set conn = Nothing  ' cached conn — do not close
    End If
End Function

'=============================================================================
' FindSqlKeyword — 定位字符串字面量/方括号标识符之外、词边界合格的 SQL 关键词
'
' 用途: SqlRangeQuery 解析 SELECT 文本中的 WHERE / ORDER BY, 避免字符串
'   字面量 ('WHERE') 或表名 ([DataWhere$]) 中的子串误命中 (R5-05 / R5-36)。
' 单引号字符串内 '' 与方括号标识符内 ]] 为转义序列, 其内部文本不参与匹配。
' 词边界: 关键词前后须为字符串首尾、空白、()、; 或 ,
'
' @Returns 关键词 1-based 起始位置; 0 = 未找到
'=============================================================================
Private Function FindSqlKeyword(ByVal sql As String, ByVal keyword As String) As Long
    Dim i As Long, kLen As Long
    Dim ch1 As String, ch2 As String
    Dim beforeCh As String, afterCh As String
    Dim inQuote As Boolean, inBracket As Boolean

    kLen = Len(keyword)
    If kLen = 0 Then Exit Function
    i = 1
    Do While i <= Len(sql)
        ch1 = Mid$(sql, i, 1)
        ch2 = ""
        If i < Len(sql) Then ch2 = Mid$(sql, i + 1, 1)
        If inQuote Then
            If ch1 = "'" Then
                If ch2 = "'" Then i = i + 2 Else inQuote = False: i = i + 1
            Else
                i = i + 1
            End If
        ElseIf inBracket Then
            If ch1 = "]" Then
                If ch2 = "]" Then i = i + 2 Else inBracket = False: i = i + 1
            Else
                i = i + 1
            End If
        ElseIf ch1 = "'" Then
            inQuote = True: i = i + 1
        ElseIf ch1 = "[" Then
            inBracket = True: i = i + 1
        Else
            If StrComp(Mid$(sql, i, kLen), keyword, vbTextCompare) = 0 Then
                beforeCh = "": If i > 1 Then beforeCh = Mid$(sql, i - 1, 1)
                afterCh = "": If i + kLen <= Len(sql) Then afterCh = Mid$(sql, i + kLen, 1)
                If IsSqlWordBoundary(beforeCh) And IsSqlWordBoundary(afterCh) Then
                    FindSqlKeyword = i
                    Exit Function
                End If
            End If
            i = i + 1
        End If
    Loop
End Function

'=============================================================================
' IsSqlWordBoundary — SQL 关键词边界字符判定 (首尾/空白/圆括号/分号/逗号)
'=============================================================================
Private Function IsSqlWordBoundary(ByVal ch As String) As Boolean
    IsSqlWordBoundary = (ch = "" Or ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf _
        Or ch = "(" Or ch = ")" Or ch = ";" Or ch = ",")
End Function

'=============================================================================
' SqlRangeQuery — 对 Range 直接查询 (无需保存工作簿)
' 支持: SELECT * (返回所有列) | WHERE (rs.Filter)
' 不支持: 指定列、ORDER BY、JOIN、GROUP BY、聚合函数
'
' ⚠️ WHERE 能力边界 (R5-26): 子句交由 ADODB Recordset.Filter 执行, 其 LIKE
'   引擎与 ACE SQL 不同 — 方括号按字面量匹配 (SqlEscapeString(forLike:=True)
'   的输出仅适用于 ACE SQL 路径), 且不支持 IN / NOT IN / BETWEEN。
' ⚠️ 单单元格 Range 输入直接原样返回, 不解析/应用 WHERE, tableAlias 仅作占位。
'=============================================================================
Public Function SqlRangeQuery( _
    ByVal sql As String, _
    ByRef rng As Range, _
    Optional ByVal tableAlias As String = "data", _
    Optional ByRef outOk As Boolean) As Variant()

    Dim rs As Object, data As Variant
    Dim nRows As Long, nCols As Long, i As Long, j As Long
    Dim result() As Variant

    outOk = False
    If rng Is Nothing Then
        Err.Raise ERR_INVALID_INPUT, "SqlUtils", "Range 不能为 Nothing。"
    End If

    On Error GoTo RangeQueryErr
    ' 多区域检查置于 handler 之后 — 保持 outOk=False 软失败契约 (与 rng.Value 失败路径一致)
    If rng.Areas.Count > 1 Then
        Err.Raise ERR_INVALID_INPUT, "SqlRangeQuery", "不支持多区域 Range，请使用单个连续区域。"
    End If
    ' 基本语法门禁: 仅支持 SELECT ... FROM ...（GARBAGE SQL 不再静默返回全表）
    Dim sqlTrim As String: sqlTrim = UCase$(Trim$(sql))
    If Left$(sqlTrim, 6) <> "SELECT" Or InStr(1, sqlTrim, " FROM ", vbTextCompare) = 0 Then
        Err.Raise ERR_INVALID_INPUT, "SqlRangeQuery", _
            "仅支持 SELECT ... FROM ... [WHERE ...] 语法: " & sql
    End If
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3  ' adUseClient
    data = rng.Value

    ' 单格输入无表头/行概念: 设计上原样返回单值, 不解析/应用 WHERE (tableAlias 忽略) — R5-59
    If rng.Count = 1 Then
        ReDim result(1 To 1, 1 To 2)
        result(1, 1) = "F1": result(1, 2) = data
        outOk = True: SqlRangeQuery = result
        Exit Function
    End If

    nRows = UBound(data, 1): nCols = UBound(data, 2)

    ' Build field names with deduplication
    ' (vbTextCompare: ADO Fields 集合名大小写不敏感, "Name"/"name" 必须去重)
    Dim colNames As Object, colName As String, colSuffix As Long
    Set colNames = CreateObject("Scripting.Dictionary")
    colNames.CompareMode = vbTextCompare
    For j = 1 To nCols
        colName = MakeSafeColumnName(CStr(data(1, j)))
        If colNames.Exists(colName) Then
            colSuffix = 2
            Do While colNames.Exists(colName & "_" & colSuffix)
                colSuffix = colSuffix + 1
            Loop
            colName = colName & "_" & colSuffix
        End If
        colNames.Add colName, True
        ' 自动检测列类型:
        '   数值子类型白名单 → adDouble (防止字符串比较陷阱 #45)
        '   Date → adDate; 数字样文本 ("01234"/长ID) 不转换, 防止前导零/精度损坏
        '   Error 单元格跳过 (不参与推断, 也不写入)
        Dim hasNum As Boolean: hasNum = False
        Dim hasDate As Boolean: hasDate = False
        Dim hasText As Boolean: hasText = False
        Dim vt As Long
        For i = 2 To nRows
            If Not IsEmpty(data(i, j)) And Not IsNull(data(i, j)) And Not IsError(data(i, j)) Then
                vt = VarType(data(i, j))
                Select Case vt
                    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                        hasNum = True
                    Case vbDate
                        hasDate = True
                    Case Else
                        hasText = True
                        Exit For
                End Select
            End If
        Next i
        If hasText Or (hasNum And hasDate) Then
            rs.Fields.Append colName, 202, 32767    ' adVarWChar, DefinedSize=32767
        ElseIf hasDate Then
            rs.Fields.Append colName, 7             ' adDate
        ElseIf hasNum Then
            rs.Fields.Append colName, 5, 32         ' adDouble, DefinedSize=32
        Else
            rs.Fields.Append colName, 202, 32767    ' 全空列 → 文本
        End If
    Next j
    rs.Open
    For i = 2 To nRows
        rs.AddNew
        For j = 1 To nCols
            If Not IsEmpty(data(i, j)) And Not IsNull(data(i, j)) And Not IsError(data(i, j)) Then
                rs.Fields(j - 1).Value = data(i, j)
            End If
        Next j
        rs.Update
    Next i

    ' Extract WHERE clause → rs.Filter (词边界识别; 跳过字符串字面量/方括号标识符,
    ' 容忍 WHERE( 与制表符; R5-05: 首个命中不合格时继续向后搜索)
    Dim wherePos As Long
    wherePos = FindSqlKeyword(sql, "WHERE")
    ' R5-36: ORDER BY 同样只在引号/标识符之外按词边界检测 (WHERE Name='ORDER BY x' 合法)
    If FindSqlKeyword(sql, "ORDER BY") > 0 Then
        Err.Raise ERR_INVALID_INPUT, "SqlRangeQuery", _
            "不支持 ORDER BY — 仅支持 SELECT * ... [WHERE ...]"
    End If
    If wherePos > 0 Then
        Dim filterStr As String
        filterStr = Trim$(Mid$(sql, wherePos + 5))
        If Len(filterStr) > 0 Then
            On Error Resume Next
            rs.Filter = filterStr
            Dim filterErr As Long: filterErr = Err.Number
            On Error GoTo RangeQueryErr
            If filterErr <> 0 Then
                Err.Raise filterErr, "SqlRangeQuery", "ADODB Recordset Filter 失败: " & filterStr
            End If
        End If
    End If

    ' Guard: filter may have returned zero rows
    If rs.EOF Then
        ReDim result(1 To 1, 1 To 1): result(1, 1) = Empty
        outOk = True: SqlRangeQuery = result: GoTo Cleanup
    End If
    If Not rs.EOF Then rs.MoveFirst
    Dim rawData As Variant
    rawData = rs.GetRows()  ' all rows (no limit)
    nRows = UBound(rawData, 2) + 1
    nCols = rs.Fields.Count

    ReDim result(1 To nRows + 1, 1 To nCols)
    For j = 0 To nCols - 1
        result(1, j + 1) = rs.Fields(j).Name
    Next j
    For i = 0 To nRows - 1
        For j = 0 To nCols - 1
            If IsNull(rawData(j, i)) Then
                result(i + 2, j + 1) = Empty
            Else
                result(i + 2, j + 1) = rawData(j, i)
            End If
        Next j
    Next i

    outOk = True: SqlRangeQuery = result
    GoTo Cleanup

RangeQueryErr:
    outOk = False
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        If Err.Number <> 0 Then Err.Clear
        Set rs = Nothing
        On Error GoTo 0
    End If
    Exit Function

Cleanup:
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        If Err.Number <> 0 Then Err.Clear
        Set rs = Nothing
        On Error GoTo 0
    End If
End Function

'=============================================================================
' 工作表函数 (UDF_SQL_*)
'=============================================================================

Public Function UDF_SQL_QUERY(ByVal sql As Variant, _
    Optional ByVal filePath As Variant = "") As Variant
    On Error GoTo EH
    Dim ok As Boolean
    UDF_SQL_QUERY = SqlExecute(sql, filePath, True, ok)
    If Not ok Then UDF_SQL_QUERY = CVErr(xlErrValue)
    Exit Function
EH: UDF_SQL_QUERY = CVErr(xlErrValue)
End Function

Public Function UDF_SQL_JOIN( _
    ByVal table1 As Variant, ByVal table2 As Variant, _
    ByVal joinOn As Variant, _
    Optional ByVal joinType As Variant = "INNER", _
    Optional ByVal selectCols As Variant = "*", _
    Optional ByVal filePath As Variant = "") As Variant
    On Error GoTo EH
    Dim ok As Boolean
    UDF_SQL_JOIN = SqlJoin(table1, table2, joinOn, joinType, selectCols, filePath, ok)
    If Not ok Then UDF_SQL_JOIN = CVErr(xlErrValue)
    Exit Function
EH: UDF_SQL_JOIN = CVErr(xlErrValue)
End Function

Public Function UDF_SQL_GROUPBY( _
    ByVal tableName As Variant, ByVal groupCols As Variant, _
    ByVal aggExprs As Variant, _
    Optional ByVal whereClause As Variant = "", _
    Optional ByVal filePath As Variant = "") As Variant
    On Error GoTo EH
    Dim ok As Boolean
    UDF_SQL_GROUPBY = SqlGroupBy(tableName, groupCols, aggExprs, whereClause, filePath, ok)
    If Not ok Then UDF_SQL_GROUPBY = CVErr(xlErrValue)
    Exit Function
EH: UDF_SQL_GROUPBY = CVErr(xlErrValue)
End Function

Public Function UDF_SQL_LIST_SHEETS( _
    Optional ByVal filePath As Variant = "") As Variant
    On Error GoTo EH
    Dim ok As Boolean
    UDF_SQL_LIST_SHEETS = SqlListSheets(filePath, ok)
    If Not ok Then UDF_SQL_LIST_SHEETS = CVErr(xlErrValue)
    Exit Function
EH: UDF_SQL_LIST_SHEETS = CVErr(xlErrValue)
End Function

Public Function UDF_SQL_LIST_COLUMNS(ByVal tableName As Variant, _
    Optional ByVal filePath As Variant = "") As Variant
    On Error GoTo EH
    Dim ok As Boolean
    UDF_SQL_LIST_COLUMNS = SqlListColumns(tableName, filePath, ok)
    If Not ok Then UDF_SQL_LIST_COLUMNS = CVErr(xlErrValue)
    Exit Function
EH: UDF_SQL_LIST_COLUMNS = CVErr(xlErrValue)
End Function

Public Function UDF_SQL_LIST_TABLES( _
    Optional ByVal filePath As Variant = "") As Variant
    On Error GoTo EH
    Dim ok As Boolean
    UDF_SQL_LIST_TABLES = SqlListTables(filePath, ok)
    If Not ok Then UDF_SQL_LIST_TABLES = CVErr(xlErrValue)
    Exit Function
EH: UDF_SQL_LIST_TABLES = CVErr(xlErrValue)
End Function

'=====================================================================
' 使用示例
'=====================================================================
' ' 基础查询
' Dim ok As Boolean, result As Variant
' result = SqlExecute("SELECT * FROM [Sheet1$] WHERE Age > 30", , , ok)
' Range("A1").Resize(UBound(result,1), UBound(result,2)).Value = result
'
' ' JOIN
' result = SqlJoin("Sales", "Regions", "t1.RegionID = t2.ID", "LEFT", , , ok)
'
' ' 分组聚合
' result = SqlGroupBy("Sales", "Region", "SUM(Amount) AS Total, COUNT(*) AS Cnt", , , ok)
'
' ' 工作表公式:
' ' =UDF_SQL_QUERY("SELECT * FROM [Sheet1$] WHERE Age > 30")
' ' =UDF_SQL_JOIN("Sales", "Regions", "t1.RegionID = t2.ID")
' ' =UDF_SQL_GROUPBY("Sales", "Region", "SUM(Amount) AS Total")
' ' =UDF_SQL_LIST_SHEETS()
' ' =UDF_SQL_LIST_COLUMNS("Sales")
' ' =UDF_SQL_LIST_TABLES()

'=====================================================================
' Test_SqlUtils — Debug.Assert test harness
' Run: Call Test_SqlUtils (any assertion failure halts VBA)
'=====================================================================
Public Sub Test_SqlUtils()
    Dim rng As Range, result As Variant, ok As Boolean
    Dim ws As Worksheet, i As Long
    Dim failed As Boolean
    Dim savedErrNum As Long, savedErrDesc As String

    On Error GoTo TestError

    '=====================================================================
    ' SqlEscapeString / NormalizeTableName — 单元断言
    '=====================================================================
    If Not (SqlEscapeString("O'Brien's") = "O''Brien''s") Then Err.Raise 5
    If Not (SqlEscapeString("a[b]", True) = "a[[]b]") Then Err.Raise 5
    If Not (SqlEscapeString("100%_x[", True) = "100[%][_]x[[]") Then Err.Raise 5
    ' R5-27: 含空格表名四种入参形式归一化一致
    If Not (NormalizeTableName("Data Sheet") = "Data Sheet$") Then Err.Raise 5
    If Not (NormalizeTableName("'Data Sheet$'") = "Data Sheet$") Then Err.Raise 5
    If Not (NormalizeTableName("Data Sheet$") = "Data Sheet$") Then Err.Raise 5
    If Not (NormalizeTableName("[Data Sheet$]") = "Data Sheet$") Then Err.Raise 5

    '=====================================================================
    ' SqlRangeQuery — SELECT *
    '=====================================================================
    ' Create a simple test range on a temporary sheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Range("A1").Value = "ID"
    ws.Range("B1").Value = "Name"
    ws.Range("C1").Value = "Score"
    ws.Range("A2").Value = 1
    ws.Range("B2").Value = "Alice"
    ws.Range("C2").Value = 90
    ws.Range("A3").Value = 2
    ws.Range("B3").Value = "Bob"
    ws.Range("C3").Value = 85
    ws.Range("A4").Value = 3
    ws.Range("B4").Value = "Charlie"
    ws.Range("C4").Value = 95
    Set rng = ws.Range("A1:C4")

    result = SqlRangeQuery("SELECT * FROM data", rng, "data", ok)
    If Not (ok = True) Then Err.Raise 5
    If Not (UBound(result, 1) = 4) Then Err.Raise 5
    If Not (UBound(result, 2) = 3) Then Err.Raise 5
    If Not (result(1, 1) = "ID") Then Err.Raise 5
    If Not (result(1, 2) = "Name") Then Err.Raise 5
    If Not (result(1, 3) = "Score") Then Err.Raise 5
    If Not (result(2, 1) = 1) Then Err.Raise 5
    If Not (result(2, 2) = "Alice") Then Err.Raise 5
    If Not (result(3, 1) = 2) Then Err.Raise 5
    If Not (result(4, 1) = 3) Then Err.Raise 5

    ' SqlRangeQuery — SELECT * with WHERE filter
    result = SqlRangeQuery("SELECT * FROM data WHERE Score > 90", rng, "data", ok)
    If Not (ok = True) Then Err.Raise 5
    If Not (UBound(result, 1) = 2) Then Err.Raise 5
    If Not (result(2, 1) = 3) Then Err.Raise 5
    If Not (result(2, 2) = "Charlie") Then Err.Raise 5

    ' SqlRangeQuery — WHERE with no matches
    result = SqlRangeQuery("SELECT * FROM data WHERE Score > 999", rng, "data", ok)
    If Not (ok = True) Then Err.Raise 5
    If Not (UBound(result, 1) = 1) Then Err.Raise 5
    If Not (IsEmpty(result(1, 1))) Then Err.Raise 5

    ' SqlRangeQuery — R5-05: 字符串字面量/表名中的 WHERE 不得吞掉过滤
    result = SqlRangeQuery("SELECT 'WHERE' AS w FROM data WHERE Score > 90", rng, "data", ok)
    If Not (ok = True) Then Err.Raise 5
    If Not (UBound(result, 1) = 2) Then Err.Raise 5
    If Not (result(2, 2) = "Charlie") Then Err.Raise 5

    result = SqlRangeQuery("SELECT * FROM [DataWhere$] WHERE Score > 90", rng, "data", ok)
    If Not (ok = True) Then Err.Raise 5
    If Not (UBound(result, 1) = 2) Then Err.Raise 5
    If Not (result(2, 2) = "Charlie") Then Err.Raise 5

    ' SqlRangeQuery — R3 目标场景保持有效: WHERE( 与制表符
    result = SqlRangeQuery("SELECT * FROM data WHERE(Score > 90)", rng, "data", ok)
    If Not (ok = True) Then Err.Raise 5
    If Not (UBound(result, 1) = 2) Then Err.Raise 5

    result = SqlRangeQuery("SELECT * FROM data" & vbTab & "WHERE" & vbTab & "Score > 90", rng, "data", ok)
    If Not (ok = True) Then Err.Raise 5
    If Not (UBound(result, 1) = 2) Then Err.Raise 5

    ' SqlRangeQuery — R5-36: WHERE 字符串字面量中的 ORDER BY 不得误判
    result = SqlRangeQuery("SELECT * FROM data WHERE Name = 'ORDER BY x'", rng, "data", ok)
    If Not (ok = True) Then Err.Raise 5
    If Not (UBound(result, 1) = 1) Then Err.Raise 5
    If Not (IsEmpty(result(1, 1))) Then Err.Raise 5

    ' SqlRangeQuery — R5-36 反例: 真正的 ORDER BY 仍须拒绝
    On Error Resume Next
    result = SqlRangeQuery("SELECT * FROM data ORDER BY Score", rng, "data", ok)
    If Err.Number = 0 And ok Then failed = True  ' 未拒绝 = 失败 (契约: outOk=False 软失败, R5-58)
    Err.Clear: On Error GoTo TestError
    If failed Then Err.Raise 5

    ' SqlRangeQuery — single cell
    Set rng = ws.Range("A1")
    result = SqlRangeQuery("SELECT * FROM data", rng, "data", ok)
    If Not (ok = True) Then Err.Raise 5
    If Not (UBound(result, 1) = 1) Then Err.Raise 5
    If Not (UBound(result, 2) = 2) Then Err.Raise 5
    If Not (result(1, 1) = "F1") Then Err.Raise 5

    ' SqlRangeQuery — Nothing Range (应报错; OERN 下 Err.Raise 会静默通过 — R5-58)
    On Error Resume Next
    result = SqlRangeQuery("SELECT *", Nothing, "data", ok)
    If Err.Number = 0 Then failed = True
    Err.Clear: On Error GoTo TestError
    If failed Then Err.Raise 5

    '=====================================================================
    ' SqlQuery — SELECT * (需要已保存的工作簿)
    '=====================================================================
    Dim wsName As String: wsName = ws.Name
    ' 仅在已保存工作簿时测试（否则 SqlGetConnection 会报错）
    If Len(ThisWorkbook.FullName) > 0 Then
        result = SqlQuery("*", wsName, , , , ok)
        If Not (ok = True) Then Err.Raise 5
        If Not (UBound(result, 2) = 3) Then Err.Raise 5
        If Not (result(1, 1) = "ID") Then Err.Raise 5

        ' SqlQuery with WHERE
        result = SqlQuery("*", wsName, "Score > 90", , , ok)
        If Not (ok = True) Then Err.Raise 5
        If Not (UBound(result, 1) = 2) Then Err.Raise 5
    End If

    '=====================================================================
    ' SqlListSheets — 列出所有工作表名
    '=====================================================================
    If Len(ThisWorkbook.FullName) > 0 Then
        result = SqlListSheets(, ok)
        If Not (ok = True) Then Err.Raise 5
        If Not (result(1, 1) = "SheetName") Then Err.Raise 5
        ' 至少包含我们刚创建的工作表
        ' (ACE OpenSchema 返回带引号与 $ 后缀的表名, 如 'Sheet1$')
        Dim found As Boolean: found = False
        For i = 2 To UBound(result, 1)
            Dim rawName As String: rawName = Replace(CStr(result(i, 1)), "'", "")
            If Right$(rawName, 1) = "$" Then rawName = Left$(rawName, Len(rawName) - 1)
            If StrComp(rawName, wsName, vbTextCompare) = 0 Then found = True: Exit For
        Next i
        If Not (found = True) Then Err.Raise 5
    End If

    '=====================================================================
    ' SqlListColumns — 列出指定表的列名
    '=====================================================================
    If Len(ThisWorkbook.FullName) > 0 Then
        result = SqlListColumns(wsName, , ok)
        If Not (ok = True) Then Err.Raise 5
        If Not (UBound(result, 1) >= 4) Then Err.Raise 5
        If Not (result(1, 1) = "ColIndex") Then Err.Raise 5
        If Not (result(1, 2) = "ColName") Then Err.Raise 5
    End If

    '=====================================================================
    ' 边界: SqlRangeQuery — 非法 SQL / 空表名
    '=====================================================================
    On Error Resume Next
    Set rng = ws.Range("A1:C4")
    result = SqlRangeQuery("GARBAGE SQL", rng, "data", ok)
    If Err.Number = 0 And ok Then failed = True  ' 未拒绝 = 失败 (R5-58)
    Err.Clear
    result = SqlRangeQuery("SELECT * FROM ", rng, "", ok)
    If Err.Number = 0 And ok Then failed = True
    Err.Clear: On Error GoTo TestError
    If failed Then Err.Raise 5

    '=====================================================================
    ' 边界: SqlListSheets — 不存在的外部文件
    '=====================================================================
    result = SqlListSheets("C:\__nonexistent_file_99999.xlsx", ok)
    If Not (ok = False Or IsArray(result)) Then Err.Raise 5

    '=====================================================================
    ' Cleanup — 断言失败/错误路径也必须删除临时表并释放连接 (R5-60)
    '=====================================================================
TestCleanup:
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Delete
        Err.Clear
    End If
    On Error GoTo 0
    Application.DisplayAlerts = True
    CloseSqlCache
    If savedErrNum <> 0 Then Err.Raise savedErrNum, "Test_SqlUtils", savedErrDesc
    Exit Sub

TestError:
    savedErrNum = Err.Number
    savedErrDesc = Err.Description
    Debug.Print "Test_SqlUtils FAILED: [" & savedErrNum & "] " & savedErrDesc
    Resume TestCleanup
End Sub