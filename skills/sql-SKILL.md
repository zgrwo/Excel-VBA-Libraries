---
name: sql
description: >
  SQL for Excel via ADODB — query Excel sheets as database tables using ACE.OLEDB
  or Jet.OLEDB providers. Use when writing SQL against Excel workbooks, debugging
  ADODB connection strings, building SELECT/JOIN/GROUP BY queries on sheet data,
  working with SqlUtils macros (SqlQuery, SqlJoin, SqlGroupBy, SqlRangeQuery),
  or troubleshooting Excel-as-database issues (type inference, header scanning,
  IMEX mode, max rows).
---

# SQL for Excel — ADODB Query Patterns

Battle-tested patterns from the SqlUtils.bas module.

## 1. Connection Strings

### 1.1 Provider Selection

| Provider | Supports | Requirement |
|----------|----------|-------------|
| `Microsoft.ACE.OLEDB.12.0` | .xlsx, .xlsm, .xls, .csv | Access Database Engine 2010+ (recommended) |
| `Microsoft.ACE.OLEDB.16.0` | .xlsx, .xlsm, .xls, .csv | Access Database Engine 2016+ |
| `Microsoft.Jet.OLEDB.4.0` | .xls only | Deprecated, 32-bit only |

### 1.2 Connection String Template

```vba
"Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & filePath & ";" & _
"Extended Properties=""Excel 12.0 Macro;HDR=YES;IMEX=1"""
```

**Critical parameters:**
- `HDR=YES` — First row is column headers. Use `HDR=NO` when querying raw sheets without headers.
- `IMEX=1` — Import mixed mode. Forces the provider to scan ALL rows (not just first 8) to determine column types. Without this, mixed-type columns (numbers + text) return NULL for the minority type.
- `Excel 12.0 Macro` — Required for .xlsm files. Use `Excel 12.0` for .xlsx files without macros.

## 2. Sheet Name Escaping

Excel sheet names with spaces or special characters must be escaped:

```vba
' In SQL queries, wrap sheet names in brackets with trailing $
"SELECT * FROM [Sheet1$]"

' For sheets with spaces or special chars
"SELECT * FROM ['My Sheet$']"

' Named ranges don't need the $ suffix
"SELECT * FROM [MyNamedRange]"
```

The `EscapeSheetName` function in SqlUtils handles this automatically:
```vba
Private Function EscapeSheetName(ByVal sheetName As String) As String
    ' 已带方括号: 校验配对后原样返回
    If Left$(sheetName, 1) = "[" Then
        If Right$(sheetName, 1) <> "]" Then
            Err.Raise ERR_INVALID_INPUT, "EscapeSheetName", "工作表名方括号不配对: " & sheetName
        End If
        EscapeSheetName = sheetName: Exit Function
    End If
    ' ] 转义为 ]] (SQL 标准方括号标识符转义)
    sheetName = Replace(sheetName, "]", "]]")
    If Right$(sheetName, 1) = "$" Then
        EscapeSheetName = "[" & sheetName & "]"
    Else
        EscapeSheetName = "[" & sheetName & "$]"
    End If
End Function
```

## 3. Type Inference Issues

### 3.1 The 8-Row Scan Problem

By default, the ACE provider scans only the first 8 rows to determine column types. If row 9+ contains text in a column that rows 1-8 had numbers, those rows return NULL.

**Fix:** Always use `IMEX=1` in the connection string extended properties. This forces a full scan.

### 3.2 Mixed-Type Columns

Even with `IMEX=1`, mixed columns return values as Text by default. To force numeric conversion for predominantly numeric columns:

```vba
' Use CDbl() / Val() in VBA post-processing
Dim v As Variant
v = rs.Fields(colIdx).Value
If IsNumeric(v) Then v = CDbl(v)
```

## 4. Query Patterns

### 4.1 Basic SELECT with WHERE

```sql
SELECT * FROM [TestData$] WHERE Temperature > 20
```

### 4.2 JOIN Between Sheets (with Aliases)

```sql
SELECT t1.*, t2.Price
FROM [Orders$] AS t1
INNER JOIN [Products$] AS t2 ON t1.ProductID = t2.ID
```

**Gotcha:** Self-joins must use aliases. `[TestData$] INNER JOIN [TestData$]` fails because the table names collide. Use `[TestData$] AS t1 INNER JOIN [TestData$] AS t2 ON t1.ID = t2.ID`.

### 4.3 GROUP BY with Aggregation

```sql
SELECT ProductCat, AVG(Temperature), COUNT(*)
FROM [TestData$]
GROUP BY ProductCat
```

### 4.4 TOP N with JOIN

```sql
SELECT TOP 5 t1.*
FROM [TestData$] AS t1
INNER JOIN [TestData$] AS t2 ON t1.ID = t2.ID
```

Note: When using `TOP N`, place it before the first column reference. `SELECT TOP N *` is valid; `SELECT * TOP N` is not.

## 5. ADODB Object Lifecycle

This project uses late binding for ADODB (`CreateObject("ADODB.Connection")`)
rather than early binding (`Tools → References → Microsoft ActiveX Data Objects`).
Late binding avoids version-lock on a specific ADO library (2.8 vs 6.1) and works
across all Windows versions without reference configuration. The trade-off is no
IntelliSense for ADODB types and ~2x slower method dispatch — negligible for the
query volumes typical in Excel workflows.

### 5.1 Connection Caching

SqlUtils caches connections per file path to avoid repeated overhead:

```vba
Private connCache As Object  ' Dictionary: filePath → Connection (module-level)

Public Function SqlGetConnection(ByVal filePath As String, ByRef outOk As Boolean) As Object
    If connCache Is Nothing Then Set connCache = CreateObject("Scripting.Dictionary")
    If connCache.Exists(filePath) Then
        Set SqlGetConnection = connCache(filePath)
        outOk = True
        Exit Function
    End If
    ' ... create new connection, add to cache
End Function
```

Call `CloseSqlCache()` to release all cached connections.

### 5.2 Cleanup Pattern

```vba
Dim conn As Object, rs As Object
Set conn = SqlGetConnection(filePath, ok)
If Not ok Then Exit Function

Set rs = conn.Execute(sql)
' Process results...
rs.Close
Set rs = Nothing
' Don't close conn — it's cached for reuse
```

**Never close cached connections** — they're shared across queries. Only `CloseSqlCache()` should close them.

## 6. OpenSchema Table Discovery

### 6.1 Listing Sheets and Named Ranges

```vba
' adSchemaTables = 20
Set rs = conn.OpenSchema(20)
Do While Not rs.EOF
    tableName = rs.Fields("TABLE_NAME").Value
    tableType = rs.Fields("TABLE_TYPE").Value  ' "TABLE" or "VIEW"
    rs.MoveNext
Loop
```

Filter for `TABLE_TYPE = "TABLE"` to exclude system tables. Sheet names appear as `'SheetName$'` (with single quotes and $ suffix in OpenSchema).

### 6.2 Listing Columns

```vba
' adSchemaColumns = 4
Set rs = conn.OpenSchema(4, Array(Empty, Empty, tableName, Empty))
Do While Not rs.EOF
    colName = rs.Fields("COLUMN_NAME").Value
    dataType = rs.Fields("DATA_TYPE").Value  ' ADO DataTypeEnum
    rs.MoveNext
Loop
```

## 7. SqlRangeQuery — In-Memory Range as Table

SqlUtils provides `SqlRangeQuery` for querying an in-memory Range without saving to a file:

```vba
Dim rng As Range: Set rng = ws.Range("A1").CurrentRegion
Dim result As Variant
result = SqlRangeQuery("SELECT Col1, AVG(Col3) FROM data GROUP BY Col1", rng, "data", ok)
```

This builds an **in-memory ADODB.Recordset** from the range data (no temporary workbook/file is created — nothing to leak on error paths), runs the query via client-side cursor filtering, and returns the result. Column names in the SQL must match the Range header row.

## 8. Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| "Could not find installable ISAM" | Missing ACE.OLEDB provider | Install Access Database Engine |
| "External table is not in the expected format" | Wrong connection string for file type | Use correct `Excel 12.0` vs `Excel 12.0 Macro` |
| All values NULL for a column | Type inference failed (mixed types) | Add `IMEX=1` to connection string |
| "Syntax error in FROM clause" | Sheet name not properly escaped | Wrap in brackets with `$` suffix |
| "Too few parameters. Expected 1" | Column name in WHERE doesn't exist | Check exact header spelling (case-insensitive but must match) |
| Self-join returns wrong results | Missing table aliases | Use `AS t1`, `AS t2` — reference as `t1.col`, `t2.col` |

## 9. Row Limit

ACE.OLEDB has a practical row limit of ~1,048,576 rows (Excel's max). For larger datasets, use a real database (SQLite, PostgreSQL) instead of Excel-as-database.

## 10. SQL Injection Prevention

⚠️ **ACE OLEDB Excel ISAM 驱动不支持参数化查询**（`ADODB.Command.Parameters` + `?` 占位符对 Excel 数据源无效）。

**防御方案**：使用 `SqlEscapeString()` 转义用户提供的字符串值：

```vba
' 安全: 转义用户输入
sql = "SELECT * FROM [Data$] WHERE Name = '" & SqlEscapeString(userInput) & "'"

' 不安全: 直接拼接用户输入
sql = "SELECT * FROM [Data$] WHERE Name = '" & userInput & "'"
```

**攻击面限制**（ACE/Jet Excel ISAM 受限的 SQL 支持降低了风险）：
- 不支持多语句（`;` 分隔）
- 不支持 DDL（`DROP TABLE`、`CREATE TABLE`）
- 不支持 SQL 注释（`--`、`/* */`）
- 部分版本支持 UNION

**建议**：
1. 所有来自单元格的用户输入在拼入 SQL 前通过 `SqlEscapeString()` 转义
2. 避免在 UDF 公式中直接使用 `&` 拼接单元格引用到 SQL 字符串中而不加转义
3. 对于 `IN` 子句中的值列表，转义每个元素

## 11. Related Skills

其他 Skill 文件及文档的加载时机见 [AGENTS.md](../../AGENTS.md) 文档路由表。
