# Excel VBA Libraries -- 用户手册

<!-- last_updated: 2026-06-17 -->

> English: [User Manual (EN)](../docs/VBA_LIB_User_Manual_EN.md) | [API Reference](api-reference.md)

## 目录

### 第一部分：数据处理

- **Chapter 1: ArrayUtils — 数组操作**
  - [Recipe 1.1 — 成绩分析](#recipe-score-analysis)
  - [Recipe 1.2 — 动态筛选](#recipe-dynamic-filter)
  - [Recipe 1.3 — 数据抽样](#recipe-data-sampling)
- **Chapter 2: DictSetUtils — 字典与集合运算**
  - [Recipe 2.1 — 集合运算](#recipe-set-operations)
  - [Recipe 2.2 — 频率统计](#recipe-frequency-stats)
  - [Recipe 2.3 — 笛卡尔积](#recipe-cartesian-product)
- **Chapter 3: PivotUtils — 数据重塑**
  - [Recipe 3.1 — 交叉表转明细](#recipe-cross-table-to-detail)
  - [Recipe 3.2 — 分组汇总](#recipe-group-summary)
  - [Recipe 3.3 — 交叉连接](#recipe-cross-join)
- **Chapter 4: SqlUtils — SQL 查询**
  - [Recipe 4.1 — 工作表查询](#recipe-worksheet-query)
  - [Recipe 4.2 — 多表关联](#recipe-multi-table-join)
  - [Recipe 4.3 — 分组聚合](#recipe-group-aggregation)

### 第二部分：数学与统计

- **Chapter 5: LinearUtils — 矩阵与线性代数**
  - [Recipe 5.1 — SVD 奇异值分解](#recipe-svd)
  - [Recipe 5.2 — 线性方程组求解](#recipe-linear-system)
  - [Recipe 5.3 — 多项式拟合](#recipe-polyfit)
- **Chapter 6: StatsUtils — 统计与分布**
  - [Recipe 6.1 — 描述统计](#recipe-descriptive-stats)
  - [Recipe 6.2 — 假设检验](#recipe-hypothesis-test)
  - [Recipe 6.3 — 相关性分析](#recipe-correlation)
- **Chapter 7: RegressUtils — 回归与方差分析**
  - [Recipe 7.1 — 因子重要性分析](#recipe-factor-importance)
  - [Recipe 7.2 — 单因素方差分析](#recipe-one-way-anova)
  - [Recipe 7.3 — 因子优化](#recipe-factor-optimization)

### 第三部分：文本与数据格式

- **Chapter 8: StringUtils — 字符串处理**
  - [Recipe 8.1 — 文本清洗](#recipe-text-cleaning)
  - [Recipe 8.2 — 编码转换](#recipe-encoding)
  - [Recipe 8.3 — 模糊匹配](#recipe-fuzzy-match)
- **Chapter 9: RegexUtils — 正则表达式**
  - [Recipe 9.1 — 正则提取](#recipe-regex-extract)
  - [Recipe 9.2 — 正则替换](#recipe-regex-replace)
  - [Recipe 9.3 — 文本分割](#recipe-text-split)
- **Chapter 10: JsonUtils — JSON 处理**
  - [Recipe 10.1 — JSON 路径提取](#recipe-json-path)
  - [Recipe 10.2 — 表格转 JSON](#recipe-table-to-json)
- **Chapter 11: XmlUtils — XML 解析**
  - [Recipe 11.1 — XPath 提取数据](#recipe-xpath-extract)
  - [Recipe 11.2 — XML 转表格](#recipe-xml-to-table)

### 第四部分：日期与 Excel/文件

- **Chapter 12: DateTimeUtils — 日期与时间**
  - [Recipe 12.1 — 日期信息提取](#recipe-date-info)
  - [Recipe 12.2 — 工作日计算](#recipe-workday)
  - [Recipe 12.3 — Unix 时间戳](#recipe-unix-timestamp)
- **Chapter 13: RangeUtils — 区域操作**
  - [Recipe 13.1 — 区域导出为 HTML](#recipe-range-export-html)
  - [Recipe 13.2 — 区域导出为 JSON](#recipe-range-export-json)
  - [Recipe 13.3 — 条件筛选](#recipe-conditional-filter)
- **Chapter 14: FileSystemUtils — 文件 I/O**
  - [Recipe 14.1 — 文本文件批量读取](#recipe-batch-file-read)
  - [Recipe 14.2 — 路径解析](#recipe-path-parse)
  - [Recipe 14.3 — 文件信息](#recipe-file-info)

### 第五部分：理化计算

- **Chapter 15: PhyChemUtils — 理化计算**
  - [Recipe 15.1 — 分子量与稀释计算](#recipe-molweight-dilution)
  - [Recipe 15.2 — 气体换算](#recipe-gas-conversion)
  - [Recipe 15.3 — 单位换算](#recipe-unit-conversion)

---

## Chapter 1: ArrayUtils -- 数组操作

数组的创建、排序、筛选、切片、聚合、搜索与转换工具集。输出数组默认 0-based；1D/2D 互转遵循 VBA Range 惯例 (1-based)。**模块**: `ArrayUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`ArrayDims`](#arraydims) | `(arr)` | 获取数组维数 (未初始化/标量返回 0) | Long |
| [`IsArray1D`](#isarray1d) | `(arr)` | 判断是否为一维数组 | Boolean |
| [`ArrayUnique`](#arrayunique) | `(arr)` | 去重，保留首次出现顺序 | Variant() 0-based |
| [`ArraySort`](#arraysort) | `(arr, [ascending])` | QuickSort 排序，O(n log n) | Variant() 0-based |
| [`ArrayFilterByValue`](#arrayfilterbyvalue) | `(arr, matchValue, [operator])` | 按值与比较运算符筛选 | Variant() 0-based |
| [`ArrayCountIf`](#arraycountif) | `(arr, matchValue, [operator])` | 统计符合条件的元素数量 | Long |
| [`ArrayConcat`](#arrayconcat) | `(arr1, arr2)` | 连接两个一维数组 | Variant() 0-based |
| [`ArraySlice`](#arrayslice) | `(arr, [start], [cnt])` | 提取子数组 (Python 风格索引) | Variant() 0-based |
| [`ArrayFlatten`](#arrayflatten) | `(arr)` | 二维数组按行展平为一维 | Variant() 0-based |
| [`ArrayTranspose1D`](#arraytranspose1d) | `(arr, [asColumn])` | 一维转二维列/行向量 (1-based) | Variant(,) 1-based |
| [`ArrayFind`](#arrayfind) | `(arr, value, [caseSensitive])` | 查找元素首个索引 (0-based) | Variant (Long/-1) |
| [`ArrayContains`](#arraycontains) | `(arr, value, [caseSensitive])` | 判断数组是否包含某值 | Variant (Boolean) |
| [`ArrayLookup`](#arraylookup) | `(lookupArray, lookupValue, lookupCol, [returnCols], [matchType])` | 内存数组版多列查找 | Variant |
| [`ArrayShuffle`](#arrayshuffle) | `(arr)` | Fisher-Yates 随机打乱 | Variant() 0-based |
| [`ArraySample`](#arraysample) | `(arr, n, [withReplacement])` | 随机抽样 (支持有/无放回) | Variant() 0-based |
| [`LinSpace`](#linspace) | `(start, endVal, n)` | 生成 n 个等间距点 | Double() 0-based |
| [`RangeFill`](#rangefill) | `(start, count, [stepSize])` | 生成等差数列 | Double() 0-based |
| [`ArrayChunk`](#arraychunk) | `(arr, size)` | 按固定大小分块，不足填充 Empty | Variant(,) 0-based |
| [`ArrayMin`](#arraymin) | `(arr)` | 数值最小值 (无有效数值返回 Empty) | Variant (Double/Empty) |
| [`ArrayMax`](#arraymax) | `(arr)` | 数值最大值 (无有效数值返回 Empty) | Variant (Double/Empty) |
| [`ArraySum`](#arraysum) | `(arr)` | 数值求和 (Kahan 补偿求和) | Variant (Double) |
| [`ArrayToString`](#arraytostring) | `(arr, [delimiter])` | 用分隔符连接为字符串 | Variant (String) |
| [`ArrayReverse`](#arrayreverse) | `(arr)` | 反转一维数组顺序 | Variant() 0-based |
| [`ArrayGetRow`](#arraygetrow) | `(arr, row)` | 提取二维数组指定行 (0-based) | Variant() 0-based |
| [`ArrayGetCol`](#arraygetcol) | `(arr, col)` | 提取二维数组指定列 (0-based) | Variant() 0-based |
| [`ArrayTranspose2D`](#arraytranspose2d) | `(arr)` | 二维数组行列转置 (1-based) | Variant(,) 1-based |
| [`ArrayEqual`](#arrayequal) | `(arr1, arr2, [caseSensitive])` | 逐元素相等比较 | Variant (Boolean) |
| [`ArrayProduct`](#arrayproduct) | `(arr)` | 数值元素连乘积 | Variant (Double) |
| [`CumSum`](#cumsum) | `(arr)` | 累积和 | Double() 0-based |
| [`ArgSort`](#argsort) | `(arr, [ascending])` | 返回排序后的索引数组 | Long() 0-based |
| [`ArrayAny`](#arrayany) | `(arr, matchValue, [operator])` | 是否存在任意元素满足条件 | Variant (Boolean) |
| [`ArrayAll`](#arrayall) | `(arr, matchValue, [operator])` | 是否所有元素都满足条件 | Variant (Boolean) |

| Recipe | Functions Used |
|--------|---------------|
| [成绩分析](#recipe-score-analysis) | `UDF_ARR_SORT`, `UDF_ARR_UNIQUE`, `UDF_ARR_COUNTIF`, `UDF_ARR_MIN`, `UDF_ARR_MAX`, `UDF_ARR_SUM` |
| [动态筛选](#recipe-dynamic-filter) | `UDF_ARR_FILTER`, `UDF_ARR_SLICE` |
| [数据抽样](#recipe-data-sampling) | `UDF_ARR_SHUFFLE`, `UDF_ARR_SAMPLE` |

### Recipe 1.1 -- 成绩分析

<a id="recipe-score-analysis"></a>

**场景**: 有一份学生成绩表，需要去重、排序、统计最高/最低/平均分、统计及格人数。

**输入**
|   | A | B |
|---|---|---|
| 1 | 姓名 | 成绩 |
| 2 | 张三 | 85 |
| 3 | 李四 | 92 |
| 4 | 王五 | 78 |
| 5 | 赵六 | 65 |
| 6 | 张三 | 85 |
| 7 | 钱七 | 91 |

`=UDF_ARR_SORT(UDF_ARR_UNIQUE(B2:B7))` → `{65, 78, 85, 91, 92}`

`=UDF_ARR_MIN(B2:B7)` → `65`

`=UDF_ARR_MAX(B2:B7)` → `92`

`=UDF_ARR_SUM(B2:B7)` → `411`

`=UDF_ARR_COUNTIF(B2:B7, 60, ">=")` → `6`

#### ArraySort

一维数组 QuickSort 排序，O(n log n)，不保证稳定排序。支持数值、字符串、日期。Null 排在最前，Error 排在最后。

**VBA Usage**
```vb
ArraySort(arr, [ascending]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组或 Range |
| ascending | Boolean | 可选. True=升序 (默认), False=降序 |

```vb
result = ArraySort(Array(3, 1, 4, 2))
' result → {1, 2, 3, 4}
```

**UDF Usage**
```
=UDF_ARR_SORT(arr, [ascending])
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Range/Array | 数据区域 |
| ascending | Boolean | 可选. 升序=TRUE (默认), 降序=FALSE |

**输入**
|   | A |
|---|---|
| 1 | 3 |
| 2 | 1 |
| 3 | 4 |
| 4 | 2 |

`=UDF_ARR_SORT(A1:A4, TRUE)` →

**输出**
|   | A |
|---|---|
| 1 | 1 |
| 2 | 2 |
| 3 | 3 |
| 4 | 4 |

#### ArrayUnique

一维数组去重，保留首次出现顺序。Null 视为重复合并，Empty 与 0 合并 (因 IsNumeric(Empty)=True)。

**VBA Usage**
```vb
ArrayUnique(arr) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组或 Range |

```vb
result = ArrayUnique(Array(1, 2, 2, 3, 1))
' result → {1, 2, 3}
```

**UDF Usage**
```
=UDF_ARR_UNIQUE(arr)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Range | 数据区域 |

#### ArrayFilterByValue

按值筛选一维数组。支持比较运算符和正则、contains。

**VBA Usage**
```vb
ArrayFilterByValue(arr, matchValue, [operator]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| matchValue | Variant | 匹配值 |
| operator | String | 可选. "=" (默认), "<", ">", "<=", ">=", "<>", "contains", "regex" |

```vb
result = ArrayFilterByValue(Array(3, 1, 4, 1, 5), 3, ">")
' result → {4, 5}
```

**UDF Usage**
```
=UDF_ARR_FILTER(arr, matchValue, [operator])
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Range | 数据区域 |
| matchValue | Value | 匹配值 |
| operator | String | 可选. 比较运算符，支持 "contains", "regex" |

#### ArrayCountIf

统计符合条件的元素数量。参数同 ArrayFilterByValue，返回计数而非数组。

**VBA Usage**
```vb
ArrayCountIf(arr, matchValue, [operator]) As Long
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| matchValue | Variant | 匹配值 |
| operator | String | 可选. 比较运算符，默认 "=" |

```vb
cnt = ArrayCountIf(Array(85, 92, 78, 65, 91), 80, ">")
' cnt → 3
```

**UDF Usage**
```
=UDF_ARR_COUNTIF(arr, matchValue, [operator])
```

<a id="arraymin"></a>
<a id="arraymax"></a>
<a id="arraysum"></a>
#### ArrayMin / ArrayMax / ArraySum

数值聚合：最小值、最大值、求和。ArraySum 使用 Kahan 补偿求和减少浮点误差。无有效数值时 ArrayMin/ArrayMax 返回 Empty，ArraySum 返回 0。

**VBA Usage**
```vb
ArrayMin(arr) As Variant
ArrayMax(arr) As Variant
ArraySum(arr) As Variant
```

```vb
minVal = ArrayMin(Array(85, 92, 78, 65, 91))    ' → 65
maxVal = ArrayMax(Array(85, 92, 78, 65, 91))    ' → 92
total  = ArraySum(Array(85, 92, 78, 65, 91))    ' → 411
```

**UDF Usage**
```
=UDF_ARR_MIN(arr)
=UDF_ARR_MAX(arr)
=UDF_ARR_SUM(arr)
```

#### ArrayConcat

连接两个一维数组 (或标量)。标量自动包装为单元素数组。

**VBA Usage**
```vb
ArrayConcat(arr1, arr2) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr1 | Variant | 第一个数组或标量 |
| arr2 | Variant | 第二个数组或标量 |

```vb
result = ArrayConcat(Array(1, 2), Array(3, 4))
' result → {1, 2, 3, 4}

result = ArrayConcat(42, Array(1, 2))
' result → {42, 1, 2}
```

**UDF Usage**
```
=UDF_ARR_CONCAT(arr1, arr2)
```

#### ArraySlice

取一维子数组。start 为 0-based 索引 (负数从末尾倒数)，cnt 为元素数 (-1=取到末尾)。

**VBA Usage**
```vb
ArraySlice(arr, [start], [cnt]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| start | Long | 可选. 起始索引 (0-based)，默认 0。负数从末尾倒数 |
| cnt | Long | 可选. 元素数量，默认 -1=取到末尾 |

```vb
result = ArraySlice(Array(10, 20, 30, 40), 1, 2)
' result → {20, 30}

result = ArraySlice(Array(10, 20, 30, 40), -1, 1)
' result → {40}
```

**UDF Usage**
```
=UDF_ARR_SLICE(arr, [start], [cnt])
```

#### ArrayFlatten

二维数组按行优先展平为一维数组。

**VBA Usage**
```vb
ArrayFlatten(arr) As Variant
```

```vb
' data 为 2x3 的二维数组: {{1,2,3},{4,5,6}}
result = ArrayFlatten(data)
' result → {1, 2, 3, 4, 5, 6}
```

**UDF Usage**
```
=UDF_ARR_FLATTEN(arr)
```

**输入**
|   | A | B | C |
|---|---|---|---|
| 1 | 1 | 2 | 3 |
| 2 | 4 | 5 | 6 |

`=UDF_ARR_FLATTEN(A1:C2)` → `{1,2,3,4,5,6}`

#### ArrayTranspose1D

一维数组转为二维列向量或行向量 (1-based 输出，兼容 VBA Range)。

**VBA Usage**
```vb
ArrayTranspose1D(arr, [asColumn]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| asColumn | Boolean | 可选. True=列向量 (默认), False=行向量 |

```vb
result = ArrayTranspose1D(Array(1, 2, 3), True)
' result → (1 To 3, 1 To 1): {{1},{2},{3}}
```

**UDF Usage**
```
=UDF_ARR_TRANSPOSE(arr, [asColumn])
```

#### ArrayFind

查找元素首个索引 (0-based 偏移)，未找到返回 -1。支持大小写敏感。

**VBA Usage**
```vb
ArrayFind(arr, value, [caseSensitive]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| value | Variant | 查找值 |
| caseSensitive | Boolean | 可选. 默认 False (不区分大小写) |

```vb
idx = ArrayFind(Array(10, 20, 30, 40), 30)
' idx → 2
```

**UDF Usage**
```
=UDF_ARR_FIND(arr, value, [caseSensitive])
```

#### ArrayContains

判断数组是否包含某值。委托给 ArrayFind，返回 Boolean。

**VBA Usage**
```vb
ArrayContains(arr, value, [caseSensitive]) As Variant
```

```vb
found = ArrayContains(Array("苹果", "香蕉", "橙子"), "香蕉")
' found → True
```

**UDF Usage**
```
=UDF_ARR_CONTAINS(arr, value, [caseSensitive])
```

#### ArrayLookup

内存数组版多列查找 (VLOOKUP 替代)。在二维数组的某列中查找值，返回指定列的数据。

**VBA Usage**
```vb
ArrayLookup(lookupArray, lookupValue, lookupCol, [returnCols], [matchType]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| lookupArray | Variant | 二维查找数组 (含标题或纯数据) |
| lookupValue | Variant | 查找值 |
| lookupCol | Long | 查找列索引 (1-based) |
| returnCols | Variant | 可选. 返回列号或列号数组。默认返回所有列 |
| matchType | Long | 可选. 0=精确匹配 (默认), 1=近似匹配 (升序) |

```vb
result = ArrayLookup(data, "张三", 1, Array(2, 3))
' result → 2D 行数组: 张三的列2和列3的值
```

**UDF Usage**
```
=UDF_ARR_LOOKUP(lookupArray, lookupValue, lookupCol, [returnCols], [matchType])
```

#### ArrayShuffle

Fisher-Yates 随机打乱一维数组，不修改原数组。

**VBA Usage**
```vb
ArrayShuffle(arr) As Variant
```

```vb
result = ArrayShuffle(Array(1, 2, 3, 4, 5))
' result → 随机排列，如 {3, 1, 5, 2, 4}
```

**UDF Usage**
```
=UDF_ARR_SHUFFLE(arr)
```

#### ArraySample

从数组中随机抽样。支持无放回 (默认) 和有放回两种模式。

**VBA Usage**
```vb
ArraySample(arr, n, [withReplacement]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| n | Long | 抽样数量 |
| withReplacement | Boolean | 可选. False=无放回 (默认), True=有放回 |

```vb
sample = ArraySample(Array(1, 2, 3, 4, 5, 6), 3)
' sample → 随机 3 个元素，如 {4, 1, 5}
```

**UDF Usage**
```
=UDF_ARR_SAMPLE(arr, n, [withReplacement])
```

#### LinSpace

生成 n 个等间距点，类似 NumPy `linspace`。

**VBA Usage**
```vb
LinSpace(start, endVal, n) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| start | Double | 起始值 |
| endVal | Double | 终止值 |
| n | Long | 点数 (n >= 2) |

```vb
result = LinSpace(0, 1, 5)
' result → {0, 0.25, 0.5, 0.75, 1}
```

**UDF Usage**
```
=UDF_ARR_LINSPACE(start, endVal, n)
```

#### RangeFill

生成等差数列。

**VBA Usage**
```vb
RangeFill(start, count, [stepSize]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| start | Double | 起始值 |
| count | Long | 元素数量 |
| stepSize | Double | 可选. 步长，默认 1 |

```vb
result = RangeFill(0, 5, 2)
' result → {0, 2, 4, 6, 8}
```

**UDF Usage**
```
=UDF_ARR_RANGEFILL(start, count, [stepSize])
```

#### ArrayChunk

将一维数组按固定大小分块为二维数组，每行一个块，不足用 Empty 填充。

**VBA Usage**
```vb
ArrayChunk(arr, size) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| size | Long | 每块大小 |

```vb
result = ArrayChunk(Array(1, 2, 3, 4, 5), 2)
' result → {{1, 2}, {3, 4}, {5, Empty}}  (3行 x 2列)
```

**UDF Usage**
```
=UDF_ARR_CHUNK(arr, size)
```

#### ArrayToString

用分隔符将数组元素连接为字符串。Null/Error/Object 安全处理。

**VBA Usage**
```vb
ArrayToString(arr, [delimiter]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| delimiter | String | 可选. 分隔符，默认 ", " |

```vb
s = ArrayToString(Array(1, 2, 3), "-")
' s → "1-2-3"
```

**UDF Usage**
```
=UDF_ARR_TOSTRING(arr, [delimiter])
```

#### ArrayReverse

反转一维数组顺序。

**VBA Usage**
```vb
ArrayReverse(arr) As Variant
```

```vb
result = ArrayReverse(Array(1, 2, 3))
' result → {3, 2, 1}
```

**UDF Usage**
```
=UDF_ARR_REVERSE(arr)
```

#### ArrayGetRow

提取二维数组指定行 (0-based 索引)。

**VBA Usage**
```vb
ArrayGetRow(arr, row) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 二维数组 |
| row | Long | 行索引 (0-based) |

```vb
result = ArrayGetRow(data, 1)  ' 第2行
```

**UDF Usage**
```
=UDF_ARR_GETROW(arr, row)
```

#### ArrayGetCol

提取二维数组指定列 (0-based 索引)。

**VBA Usage**
```vb
ArrayGetCol(arr, col) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 二维数组 |
| col | Long | 列索引 (0-based) |

```vb
result = ArrayGetCol(data, 0)  ' 第1列
```

**UDF Usage**
```
=UDF_ARR_GETCOL(arr, col)
```

#### ArrayTranspose2D

二维数组行列转置 (1-based 输出)。一维输入委托给 ArrayTranspose1D。

**VBA Usage**
```vb
ArrayTranspose2D(arr) As Variant
```

**UDF Usage**
```
=UDF_ARR_TRANSPOSE2D(arr)
```

**输入**
|   | A | B | C |
|---|---|---|---|
| 1 | 1 | 2 | 3 |
| 2 | 4 | 5 | 6 |

`=UDF_ARR_TRANSPOSE2D(A1:C2)` →

**输出**
|   | A | B |
|---|---|---|
| 1 | 1 | 4 |
| 2 | 2 | 5 |
| 3 | 3 | 6 |

#### ArrayEqual

逐元素比较两个一维数组是否相等。支持标量比较。

**VBA Usage**
```vb
ArrayEqual(arr1, arr2, [caseSensitive]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr1 | Variant | 第一个数组或标量 |
| arr2 | Variant | 第二个数组或标量 |
| caseSensitive | Boolean | 可选. 默认 False |

```vb
eq = ArrayEqual(Array(1, 2, 3), Array(1, 2, 3))
' eq → True

eq = ArrayEqual(Array(1, 2), Array(1, 2, 3))
' eq → False
```

**UDF Usage**
```
=UDF_ARR_EQUAL(arr1, arr2, [caseSensitive])
```

#### ArrayProduct

数值元素连乘积。无有效数值返回 0 (与 Excel PRODUCT 兼容)。

**VBA Usage**
```vb
ArrayProduct(arr) As Variant
```

```vb
prod = ArrayProduct(Array(2, 3, 4))
' prod → 24
```

**UDF Usage**
```
=UDF_ARR_PRODUCT(arr)
```

#### CumSum

累积和。返回与输入等长的数组，第 i 个元素 = arr(0..i) 之和。非数值元素保持累和。

**VBA Usage**
```vb
CumSum(arr) As Variant
```

```vb
result = CumSum(Array(3, 1, 4))
' result → {3, 4, 8}
```

**UDF Usage**
```
=UDF_ARR_CUMSUM(arr)
```

#### ArgSort

返回排序后的索引数组 (类似 NumPy `argsort`)。result[0] 是最小值的原始索引。

**VBA Usage**
```vb
ArgSort(arr, [ascending]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| ascending | Boolean | 可选. 默认 True |

```vb
result = ArgSort(Array(30, 10, 20))
' result → {1, 2, 0}  (值10在索引1, 20在索引2, 30在索引0)
```

**UDF Usage**
```
=UDF_ARR_ARGSORT(arr, [ascending])
```

<a id="arrayany"></a>
<a id="arrayall"></a>
#### ArrayAny / ArrayAll

判断是否存在任意/所有元素满足条件。

**VBA Usage**
```vb
ArrayAny(arr, matchValue, [operator]) As Variant
ArrayAll(arr, matchValue, [operator]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组或标量 |
| matchValue | Variant | 匹配值 |
| operator | String | 可选. 比较运算符，默认 "=" |

```vb
hasLarge = ArrayAny(Array(1, 2, 10), 5, ">")    ' → True
allPositive = ArrayAll(Array(2, 4, 6), 1, ">")   ' → True
```

**UDF Usage**
```
=UDF_ARR_ANY(arr, matchValue, [operator])
=UDF_ARR_ALL(arr, matchValue, [operator])
```

#### ArrayDims

获取数组维数。未初始化数组/标量返回 0。

**VBA Usage**
```vb
ArrayDims(arr) As Long
```

```vb
d = ArrayDims(Array(1, 2, 3))   ' d → 1
d = ArrayDims(Empty)             ' d → 0
```

#### IsArray1D

判断是否为一维数组。未初始化数组返回 False。

**VBA Usage**
```vb
IsArray1D(arr) As Boolean
```

```vb
b = IsArray1D(Array(1, 2))      ' b → True
b = IsArray1D("scalar")          ' b → False
```

### Recipe 1.2 -- 动态筛选

<a id="recipe-dynamic-filter"></a>

**场景**: 从产品列表中筛选价格在指定区间的产品，并取前 N 个。

**输入**
|   | A | B |
|---|---|---|
| 1 | 产品 | 价格 |
| 2 | 笔记本电脑 | 5999 |
| 3 | 鼠标 | 199 |
| 4 | 键盘 | 499 |
| 5 | 显示器 | 2299 |
| 6 | 耳机 | 299 |
| 7 | 平板 | 3499 |

`=UDF_ARR_SLICE(UDF_ARR_FILTER(B2:B7, 3000, "<"), 0, 3)` → `{199, 499, 2299}`

### Recipe 1.3 -- 数据抽样

<a id="recipe-data-sampling"></a>

**场景**: 从学生名单中随机抽取 3 名参加活动，不重复抽样。

**输入**
|   | A |
|---|---|
| 1 | 姓名 |
| 2 | 张三 |
| 3 | 李四 |
| 4 | 王五 |
| 5 | 赵六 |
| 6 | 钱七 |
| 7 | 孙八 |

`=UDF_ARR_SHUFFLE(A2:A7)` → 随机排序结果

`=UDF_ARR_SAMPLE(A2:A7, 3)` → 随机 3 个不重复姓名

---

## Chapter 2: DictSetUtils -- 字典与集合运算

字典创建、合并、排序、筛选与集合运算 (并/交/差/对称差/子集/笛卡尔积)。返回 Object 的函数标记 **仅 VBA**。**模块**: `DictSetUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`DictMerge`](#dictmerge) | `(dict1, dict2, [overwrite])` | 合并两个字典 | Object **仅 VBA** |
| [`DictMergeSum`](#dictmergesum) | `(dict1, dict2)` | 合并字典，同键数值求和 | Object **仅 VBA** |
| [`DictInvert`](#dictinvert) | `(dict)` | 键值互换，跳过非法值 | Object **仅 VBA** |
| [`DictKeys`](#dictkeys) | `(dict)` | 返回所有键的数组 | Variant() |
| [`DictValues`](#dictvalues) | `(dict)` | 返回所有值的数组 | Variant() |
| [`DictTo2DArray`](#dictto2darray) | `(dict)` | 字典转为 2D 键值对数组 | Variant(,) 1-based |
| [`ArrayToDict`](#arraytodict) | `(arr, [keyFn])` | 一维数组转为字典 | Object **仅 VBA** |
| [`DictFrom2DArray`](#dictfrom2darray) | `(arr, [colKey], [colValue])` | 2D 数组的两列转为字典 | Object **仅 VBA** |
| [`CountFrequency`](#countfrequency) | `(arr)` | 频数统计，返回字典 | Object **仅 VBA** |
| [`GroupCount`](#groupcount) | `(arr)` | 频数统计，直接输出 2D 数组 | Variant(,) 1-based |
| [`DictPick`](#dictpick) | `(dict, keys)` | 按键列表提取子字典 | Object **仅 VBA** |
| [`DictCount`](#dictcount) | `(dict)` | 字典条目数 | Long |
| [`SetUnion`](#setunion) | `(arr1, arr2)` | 集合并集 | Variant() 0-based |
| [`SetIntersect`](#setintersect) | `(arr1, arr2)` | 集合交集 | Variant() 0-based |
| [`SetDifference`](#setdifference) | `(arr1, arr2)` | 集合差集 (A - B) | Variant() 0-based |
| [`SetSymDifference`](#setsymdifference) | `(arr1, arr2)` | 集合对称差 (A Δ B) | Variant() 0-based |
| [`SetIsSubset`](#setissubset) | `(arr1, arr2)` | 判断 arr1 是否为 arr2 的子集 | Boolean |
| [`SetEqual`](#setequal) | `(arr1, arr2)` | 判断两集合是否相等 | Boolean |
| [`DictFilterByValue`](#dictfilterbyvalue) | `(dict, matchValue, [operator])` | 按值条件筛选子字典 | Object **仅 VBA** |
| [`DictSortByKey`](#dictsortbykey) | `(dict, [ascending])` | 按键排序，返回 2D 数组 | Variant(,) 1-based |
| [`DictSortByValue`](#dictsortbyvalue) | `(dict, [ascending])` | 按值排序，返回 2D 数组 | Variant(,) 1-based |
| [`DictGetDefault`](#dictgetdefault) | `(dict, key, [defaultValue])` | 安全取值，键不存在返回默认 | Variant |
| [`DictRenameKey`](#dictrenamekey) | `(dict, oldKey, newKey)` | 重命名键，返回新字典 | Object **仅 VBA** |
| [`DictRemoveKeys`](#dictremovekeys) | `(dict, keys)` | 批量删除键，返回新字典 | Object **仅 VBA** |
| [`DictClone`](#dictclone) | `(dict)` | 浅克隆字典 | Object **仅 VBA** |
| [`DictIsEmpty`](#dictisempty) | `(dict)` | 判断字典是否为空 | Boolean |
| [`DictTopN`](#dicttopn) | `(dict, n, [ascending])` | 按值取前 N 条，返回 2D 数组 | Variant(,) 1-based |
| [`SetCartesianProduct`](#setcartesianproduct) | `(arrA, arrB)` | 笛卡尔积，返回 2D 配对数组 | Variant(,) 1-based |

| Recipe | Functions Used |
|--------|---------------|
| [集合运算](#recipe-set-operations) | `UDF_DICT_UNION`, `UDF_DICT_INTERSECT`, `UDF_DICT_DIFFERENCE` |
| [频率统计](#recipe-frequency-stats) | `UDF_DICT_GROUPCOUNT`, `CountFrequency` |
| [笛卡尔积](#recipe-cartesian-product) | `UDF_DICT_CARTESIAN` |

### Recipe 2.1 -- 集合运算

<a id="recipe-set-operations"></a>

**场景**: 两个班级的学生名单，求并集 (全量)、交集 (共同学生)、差集 (仅 A 班)。

**输入**
|   | A | B |
|---|---|---|
| 1 | 1班 | 2班 |
| 2 | 张三 | 李四 |
| 3 | 王五 | 张三 |
| 4 | 赵六 | 钱七 |

`=UDF_DICT_UNION(A2:A4, B2:B4)` → `{张三, 王五, 赵六, 李四, 钱七}`

`=UDF_DICT_INTERSECT(A2:A4, B2:B4)` → `{张三}`

`=UDF_DICT_DIFFERENCE(A2:A4, B2:B4)` → `{王五, 赵六}`

#### SetUnion

集合并集。返回两个一维数组所有不重复元素的集合。

**VBA Usage**
```vb
SetUnion(arr1, arr2) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr1 | Variant | 第一个一维数组 |
| arr2 | Variant | 第二个一维数组 |

```vb
result = SetUnion(Array(1, 2, 3), Array(3, 4, 5))
' result → {1, 2, 3, 4, 5}
```

**UDF Usage**
```
=UDF_DICT_UNION(arr1, arr2)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr1 | Range | 第一个区域 |
| arr2 | Range | 第二个区域 |

#### SetIntersect

集合交集。返回同时存在于两个数组的元素。

**VBA Usage**
```vb
SetIntersect(arr1, arr2) As Variant
```

```vb
result = SetIntersect(Array(1, 2, 3), Array(2, 3, 4))
' result → {2, 3}
```

**UDF Usage**
```
=UDF_DICT_INTERSECT(arr1, arr2)
```

#### SetDifference

集合差集 (A - B)。返回在 arr1 中但不在 arr2 中的元素。

**VBA Usage**
```vb
SetDifference(arr1, arr2) As Variant
```

```vb
result = SetDifference(Array(1, 2, 3), Array(2, 4))
' result → {1, 3}
```

**UDF Usage**
```
=UDF_DICT_DIFFERENCE(arr1, arr2)
```

#### SetSymDifference

集合对称差 (A Δ B)。返回仅在一边出现的元素。

**VBA Usage**
```vb
SetSymDifference(arr1, arr2) As Variant
```

```vb
result = SetSymDifference(Array(1, 2, 3), Array(2, 4))
' result → {1, 3, 4}
```

**UDF Usage**
```
=UDF_DICT_SYM_DIFF(arr1, arr2)
```

#### SetIsSubset

判断 arr1 是否为 arr2 的子集。空集视为任意集合的子集。

**VBA Usage**
```vb
SetIsSubset(arr1, arr2) As Boolean
```

```vb
isSub = SetIsSubset(Array(1, 2), Array(1, 2, 3))
' isSub → True
```

**UDF Usage**
```
=UDF_DICT_ISSUBSET(arr1, arr2)
```

#### SetEqual

判断两集合是否相等 (元素完全相同)。

**VBA Usage**
```vb
SetEqual(arr1, arr2) As Boolean
```

```vb
eq = SetEqual(Array(1, 2, 3), Array(3, 2, 1))
' eq → True (顺序无关)
```

**UDF Usage**
```
=UDF_DICT_ISEQUAL(arr1, arr2)
```

### Recipe 2.2 -- 频率统计

<a id="recipe-frequency-stats"></a>

**场景**: 统计销售记录中各地区的出现频次。

**输入**
|   | A |
|---|---|
| 1 | 地区 |
| 2 | 北京 |
| 3 | 上海 |
| 4 | 北京 |
| 5 | 广州 |
| 6 | 上海 |
| 7 | 北京 |

`=UDF_DICT_GROUPCOUNT(A2:A7)` →

**输出**
|   | A | B |
|---|---|---|
| 1 | 北京 | 3 |
| 2 | 上海 | 2 |
| 3 | 广州 | 1 |

#### GroupCount

频数统计，直接输出 2D 数组 (key, count)。跳过 Error 和 Null 值。

**VBA Usage**
```vb
GroupCount(arr) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |

```vb
result = GroupCount(Array("A", "B", "A", "C", "B", "A"))
' result → {{"A", 3}, {"B", 2}, {"C", 1}}
```

**UDF Usage**
```
=UDF_DICT_GROUPCOUNT(arr)
```

#### CountFrequency

频数统计返回字典 (key -> count)。**仅 VBA**，返回 Scripting.Dictionary。

**VBA Usage**
```vb
CountFrequency(arr) As Object
```

```vb
Dim freq As Object
Set freq = CountFrequency(Array("A", "B", "A", "C"))
' freq("A") → 2, freq("B") → 1, freq("C") → 1
```

#### DictMerge

合并两个字典。overwrite=False 时重复键保留 dict1 的值；overwrite=True 时用 dict2 覆盖。

**VBA Usage**
```vb
DictMerge(dict1, dict2, [overwrite]) As Object
```

| 参数 | 类型 | 说明 |
|------|------|------|
| dict1 | Object | 第一个字典 |
| dict2 | Object | 第二个字典 |
| overwrite | Boolean | 可选. 默认 False (保留 dict1) |

```vb
Dim merged As Object
Set merged = DictMerge(dict1, dict2, True)
```

#### DictMergeSum

合并两个字典，同键数值求和。非数值键保留 dict2 的值。

**VBA Usage**
```vb
DictMergeSum(dict1, dict2) As Object
```

#### DictInvert

键值互换。跳过 Object/Error/Null/Array 类型的值 (无法作为字典键)。**仅 VBA**。

**VBA Usage**
```vb
DictInvert(dict) As Object
```

```vb
Dim inverted As Object
Set inverted = DictInvert(dict)
```

<a id="dictkeys"></a>
<a id="dictvalues"></a>
#### DictKeys / DictValues

返回字典所有键/值的 Variant 数组。

**VBA Usage**
```vb
DictKeys(dict) As Variant
DictValues(dict) As Variant
```

```vb
keys = DictKeys(dict)
vals = DictValues(dict)
```

#### DictTo2DArray

字典转为二维数组 (1 To n, 1 To 2)，第 1 列为键，第 2 列为值。

**VBA Usage**
```vb
DictTo2DArray(dict) As Variant
```

#### ArrayToDict

一维数组转为字典。keyFn="value" (默认) 以元素值为键；keyFn="index" 以索引为键。**仅 VBA**。

**VBA Usage**
```vb
ArrayToDict(arr, [keyFn]) As Object
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 一维数组 |
| keyFn | String | 可选. "value" (默认) 或 "index" |

#### DictFrom2DArray

二维数组的某两列转为字典 (colKey -> colValue)。1D 数组作为输入时自动以索引为键。**仅 VBA**。

**VBA Usage**
```vb
DictFrom2DArray(arr, [colKey], [colValue]) As Object
```

| 参数 | 类型 | 说明 |
|------|------|------|
| arr | Variant | 二维数组 |
| colKey | Long | 可选. 键列 (1-based)，默认 1 |
| colValue | Long | 可选. 值列 (1-based)，默认 2 |

#### DictPick

按指定键列表提取子字典。**仅 VBA**。

**VBA Usage**
```vb
DictPick(dict, keys) As Object
```

#### DictCount

字典条目数。Nothing 返回 0。

**VBA Usage**
```vb
DictCount(dict) As Long
```

#### DictFilterByValue

按值条件筛选子字典。支持 "=", "<", ">", "<=", ">=", "<>", "contains", "regex"。**仅 VBA**。

**VBA Usage**
```vb
DictFilterByValue(dict, matchValue, [operator]) As Object
```

<a id="dictsortbykey"></a>
<a id="dictsortbyvalue"></a>
#### DictSortByKey / DictSortByValue

按键/值排序，返回 2D 数组 (1 To n, 1 To 2)。

**VBA Usage**
```vb
DictSortByKey(dict, [ascending]) As Variant
DictSortByValue(dict, [ascending]) As Variant
```

```vb
sorted = DictSortByValue(dict, False)  ' 按值降序
' sorted(1, 1) = 最大值的键, sorted(1, 2) = 最大值
```

#### DictGetDefault

安全取值。键不存在时返回默认值，未提供默认值时返回 Empty。

**VBA Usage**
```vb
DictGetDefault(dict, key, [defaultValue]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| dict | Object | 字典 |
| key | Variant | 要查找的键 |
| defaultValue | Variant | 可选. 键不存在时的默认值 |

```vb
val = DictGetDefault(dict, "missing", 0)
' val → 0 (键不存在)
```

#### DictRenameKey

重命名字典中的键，返回新字典 (原字典不变)。新键已存在则跳过。**仅 VBA**。

**VBA Usage**
```vb
DictRenameKey(dict, oldKey, newKey) As Object
```

#### DictRemoveKeys

批量删除指定键，返回新字典 (原字典不变)。**仅 VBA**。

**VBA Usage**
```vb
DictRemoveKeys(dict, keys) As Object
```

#### DictClone

浅克隆字典。值如果是对象引用，克隆与原字典指向同一对象。**仅 VBA**。

**VBA Usage**
```vb
DictClone(dict) As Object
```

#### DictIsEmpty

判断字典是否为 Nothing 或条目数为 0。

**VBA Usage**
```vb
DictIsEmpty(dict) As Boolean
```

#### DictTopN

按值取前 N 条 (默认降序：ascending=False 取最大值前 N)，返回 2D 数组。

**VBA Usage**
```vb
DictTopN(dict, n, [ascending]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| dict | Object | 字典 |
| n | Long | 返回条目数 |
| ascending | Boolean | 可选. 默认 False (降序，最大值在前) |

```vb
top = DictTopN(dict, 3, False)
' 返回按值降序的前 3 条
```

### Recipe 2.3 -- 笛卡尔积

<a id="recipe-cartesian-product"></a>

**场景**: 生成尺寸与颜色的所有组合。

**输入**
|   | A | B |
|---|---|---|
| 1 | 尺寸 | 颜色 |
| 2 | S | 红 |
| 3 | M | 蓝 |
| 4 | L | |

`=UDF_DICT_CARTESIAN(A2:A4, B2:B3)` →

**输出**
|   | A | B |
|---|---|---|
| 1 | S | 红 |
| 2 | S | 蓝 |
| 3 | M | 红 |
| 4 | M | 蓝 |
| 5 | L | 红 |
| 6 | L | 蓝 |

#### SetCartesianProduct

两个一维数组的笛卡尔积，返回 2D 数组 (1 To nA*nB, 1 To 2)。

**VBA Usage**
```vb
SetCartesianProduct(arrA, arrB) As Variant
```

**UDF Usage**
```
=UDF_DICT_CARTESIAN(arr1, arr2)
```

---

## Chapter 3: PivotUtils -- 数据重塑

数据透视/逆透视、分组聚合、表格筛选转置与交叉连接。标注 `(Sub)` 的函数直接写入工作表，无返回值。**模块**: `PivotUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`RawConversion`](#rawconversion) | `(srcRange, valueColRef, colDimRef, rowDimRef, [keepBlank], [sortLabels])` | 交叉表透视 (长表至宽表) | Variant(,) |
| [`RawConversionToRange`](#rawconversiontorange) | `(... destCell, ...)` (Sub) | 透视结果直接输出到工作表 | 无 |
| [`Unpivot`](#unpivot) | `(rng, valueCols, nameCol, valCol, destCell, [idColIndices])` (Sub) | 逆透视 (宽表至长表) | 无 |
| [`GroupBy`](#groupby) | `(rng, groupCol, aggCol, [aggFunc], [destCell])` | 分组聚合 (SUM/COUNT/AVG/MIN/MAX) | Variant(,) |
| [`SplitColumnToRows`](#splitcolumntorows) | `(rng, colIdx, delimiter, destCell)` (Sub) | 按分隔符拆分某列到多行 | 无 |
| [`MergeColumns`](#mergecolumns) | `(rng, colIndices, delimiter, newColName, destCell, [dropOrigCols])` (Sub) | 合并多列为一列 | 无 |
| [`FilterTable`](#filtertable) | `(rng, colIdx, op, value, destCell)` (Sub) | 按条件筛选表格 (含标题行) | 无 |
| [`TransposeTable`](#transposetable) | `(rng, destCell)` (Sub) | 表格行列转置 | 无 |
| [`VLookupArray`](#vlookuparray) | `(dataArray, lookupValue, lookupCol, [returnCol])` | 数组版 VLOOKUP | Variant |
| [`CrossJoin`](#crossjoin) | `(rng1, rng2)` | 交叉连接两个表格 (笛卡尔积) | Variant(,) |

| Recipe | Functions Used |
|--------|---------------|
| [交叉表转明细](#recipe-cross-table-to-detail) | `UDF_PIVOT_CONVERT` |
| [分组汇总](#recipe-group-summary) | `UDF_PIVOT_GROUPBY` |
| [交叉连接](#recipe-cross-join) | `UDF_PIVOT_CROSSJOIN` |

### Recipe 3.1 -- 交叉表转明细

<a id="recipe-cross-table-to-detail"></a>

**场景**: 将销售原始数据按产品和月份透视成交叉表，值列显示金额。

**输入**
|   | A | B | C |
|---|---|---|---|
| 1 | 产品 | 月份 | 金额 |
| 2 | 苹果 | 1月 | 100 |
| 3 | 苹果 | 2月 | 150 |
| 4 | 香蕉 | 1月 | 80 |
| 5 | 香蕉 | 1月 | 120 |
| 6 | 苹果 | 2月 | 200 |

`=UDF_PIVOT_CONVERT(A1:C6, "金额", "月份", "产品")` →

**输出**
|   | A | B | C | D | E |
|---|---|---|---|---|---|---|
| 1 | 产品 | 1月A | 1月B | 2月A | 2月B |
| 2 | 苹果 | 100 | | 150 | 200 |
| 3 | 香蕉 | 80 | 120 | | |

> 列标题后缀 A/B 表示同一交叉点有多条记录时按出现顺序展开。

#### RawConversion

数据透视 (长表 -> 交叉表)。根据行维度、列维度、值列重新排列表格。

**VBA Usage**
```vb
RawConversion(srcRange, valueColRef, colDimRef, rowDimRef, [keepBlank], [sortLabels]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| srcRange | Range | 源数据区域 (含标题行) |
| valueColRef | Variant | 值填充列的标题文本或单元格引用 |
| colDimRef | Variant | 列维度列的标题文本或单元格引用 |
| rowDimRef | Variant | 行维度列的标题文本或单元格引用 |
| keepBlank | Boolean | 可选. 空白填充为 "" (True, 默认) 或 Empty (False) |
| sortLabels | Boolean | 可选. 是否字母升序排序行列标签，默认 False |

```vb
result = RawConversion(Range("A1:C100"), "金额", "月份", "产品", True, True)
```

**UDF Usage**
```
=UDF_PIVOT_CONVERT(srcRange, valueColRef, colDimRef, rowDimRef, [keepBlank], [sortLabels])
```

#### RawConversionToRange

将透视结果直接写入指定工作表位置。参数与 RawConversion 相同，增加 destCell 指定输出起始单元格。

**VBA Usage**
```vb
RawConversionToRange srcRange, valueColRef, colDimRef, rowDimRef, destCell, [keepBlank], [sortLabels]
```

```vb
RawConversionToRange Range("A1:C100"), "金额", "月份", "产品", Range("E1"), True, True
```

#### Unpivot

逆透视 (宽表 -> 长表)。将多列值列转为属性名/属性值两列，保留 ID 列。

**VBA Usage**
```vb
Unpivot rng, valueCols, nameCol, valCol, destCell, [idColIndices]
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rng | Range | 源区域 (含标题行) |
| valueCols | Variant | 要逆透视的值列索引 (单数或数组) |
| nameCol | String | 属性名列的标题文本 |
| valCol | String | 属性值列的标题文本 |
| destCell | Range | 输出起始单元格 |
| idColIndices | Variant | 可选. 保持不变的 ID 列索引 (默认: 非值列的其他列) |

### Recipe 3.2 -- 分组汇总

<a id="recipe-group-summary"></a>

**场景**: 按地区汇总销售额，计算总和和平均值。

**输入**
|   | A | B |
|---|---|---|
| 1 | 地区 | 销售额 |
| 2 | 北京 | 1000 |
| 3 | 上海 | 800 |
| 4 | 北京 | 1200 |
| 5 | 广州 | 600 |
| 6 | 上海 | 900 |

`=UDF_PIVOT_GROUPBY(A1:B6, 1, 2, "SUM")` →

**输出**
|   | A | B |
|---|---|---|
| 1 | 地区 | SUM(销售额) |
| 2 | 北京 | 2200 |
| 3 | 上海 | 1700 |
| 4 | 广州 | 600 |

`=UDF_PIVOT_GROUPBY(A1:B6, 1, 2, "AVG")` →

**输出**
|   | A | B |
|---|---|---|
| 1 | 地区 | AVG(销售额) |
| 2 | 北京 | 1100 |
| 3 | 上海 | 850 |
| 4 | 广州 | 600 |

#### GroupBy

分组聚合。支持 SUM/COUNT/AVG/MIN/MAX。跳过 Null/Error/非数值。

**VBA Usage**
```vb
GroupBy(rng, groupCol, aggCol, [aggFunc], [destCell]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rng | Range | 源区域 (含标题行) |
| groupCol | Long | 分组列索引 (1-based) |
| aggCol | Long | 聚合列索引 (1-based) |
| aggFunc | String | 可选. "SUM" (默认), "COUNT", "AVG", "MIN", "MAX" |
| destCell | Range | 可选. 若提供则同时写入工作表 |

```vb
result = GroupBy(Range("A1:B100"), 1, 2, "AVG")
```

**UDF Usage**
```
=UDF_PIVOT_GROUPBY(rng, groupCol, aggCol, [aggFunc])
```

#### VLookupArray

数组版 VLOOKUP。在 Range 或二维数组中按列查找，返回匹配行指定列的值。大小写不敏感。

**VBA Usage**
```vb
VLookupArray(dataArray, lookupValue, lookupCol, [returnCol]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| dataArray | Variant | 数据区域 (Range) 或二维数组 |
| lookupValue | Variant | 查找值 |
| lookupCol | Long | 查找列索引 (1-based) |
| returnCol | Long | 可选. 返回列索引 (1-based)，默认 2 |

```vb
result = VLookupArray(Range("A1:B10"), "张三", 1, 2)
```

**UDF Usage**
```
=UDF_PIVOT_VLOOKUP(dataArray, lookupValue, lookupCol, [returnCol])
```

#### SplitColumnToRows

按分隔符拆分某列到多行。拆分后的元素各自占一行，其他列复制原值。

**VBA Usage**
```vb
SplitColumnToRows rng, colIdx, delimiter, destCell
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rng | Range | 源区域 (含标题行) |
| colIdx | Long | 要拆分的列索引 (1-based) |
| delimiter | String | 分隔符 |
| destCell | Range | 输出起始单元格 |

#### MergeColumns

合并多列为一列。支持保留或删除原始列。

**VBA Usage**
```vb
MergeColumns rng, colIndices, delimiter, newColName, destCell, [dropOrigCols]
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rng | Range | 源区域 |
| colIndices | Variant | 要合并的列索引 (单数或数组) |
| delimiter | String | 分隔符 |
| newColName | String | 新列标题 |
| destCell | Range | 输出起始单元格 |
| dropOrigCols | Boolean | 可选. 是否删除原始列，默认 False |

#### FilterTable

按条件筛选表格 (保留标题行)。支持 "=", "<", ">", "<=", ">=", "<>", "contains"。

**VBA Usage**
```vb
FilterTable rng, colIdx, op, value, destCell
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rng | Range | 源区域 (含标题行) |
| colIdx | Long | 筛选列索引 (1-based) |
| op | String | 比较运算符 |
| value | Variant | 比较值 |
| destCell | Range | 输出起始单元格 |

```vb
FilterTable Range("A1:C100"), 3, ">", 60, Range("E1")
```

#### TransposeTable

表格行列转置 (第 1 行与第 1 列交换)。

**VBA Usage**
```vb
TransposeTable rng, destCell
```

### Recipe 3.3 -- 交叉连接

<a id="recipe-cross-join"></a>

**场景**: 生成学生与科目的全部组合 (排课表)。

**输入**
|   | A | B |
|---|---|---|
| 1 | 学生 | 科目 |
| 2 | 张三 | 数学 |
| 3 | 李四 | 英语 |

`=UDF_PIVOT_CROSSJOIN(A1:A3, B1:B3)` →

**输出**
|   | A | B |
|---|---|---|
| 1 | 学生 | 科目 |
| 2 | 张三 | 数学 |
| 3 | 张三 | 英语 |
| 4 | 李四 | 数学 |
| 5 | 李四 | 英语 |

#### CrossJoin

两个表格的交叉连接 (笛卡尔积)。输出行数 = 1 + (rng1数据行) * (rng2数据行)。

**VBA Usage**
```vb
CrossJoin(rng1, rng2) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rng1 | Range | 第一个区域 (含标题行) |
| rng2 | Range | 第二个区域 (含标题行) |

**UDF Usage**
```
=UDF_PIVOT_CROSSJOIN(rng1, rng2)
```

> **已知限制**: 作为工作表公式使用时溢出区域可能被截断为输入行数；完整笛卡尔积请在 VBA 中直接调用 `CrossJoin`。

---

## Chapter 4: SqlUtils -- SQL 查询

通过 ADODB (ACE/Jet OLEDB) 将 Excel 工作表当作数据库表进行 SQL 查询。支持 SELECT/JOIN/GROUP BY/WHERE/ORDER BY。**模块**: `SqlUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`CloseSqlCache`](#closesqlcache) | `()` (Sub) | 释放缓存的数据库连接 | 无 |
| [`SqlGetConnection`](#sqlgetconnection) | `([filePath], [outOk])` | 获取 ADODB 连接 (缓存复用) | Object **仅 VBA** |
| [`SqlExecute`](#sqlexecute) | `(sql, [filePath], [includeHeader], outOk, [outErrorMsg])` | 执行任意 SQL，返回 2D 数组 | Variant() **仅 VBA** |
| [`SqlQuery`](#sqlquery) | `(selectClause, fromClause, [whereClause], [orderByClause], [filePath], [outOk])` | SELECT 查询 (分句传入) | Variant() |
| [`SqlJoin`](#sqljoin) | `(table1, table2, joinOn, [joinType], [selectCols], [filePath], [outOk])` | 两表 JOIN | Variant() |
| [`SqlGroupBy`](#sqlgroupby) | `(tableName, groupCols, aggExprs, [whereClause], [filePath], [outOk])` | 分组聚合 | Variant() |
| [`SqlListSheets`](#sqllistsheets) | `([filePath], [outOk])` | 列出工作簿中的所有工作表名 | Variant() |
| [`SqlListColumns`](#sqllistcolumns) | `(tableName, [filePath], [outOk])` | 列出指定表的列名和序号 | Variant() |
| [`SqlListTables`](#sqllisttables) | `([filePath], [outOk])` | 列出所有可用数据源 (表/命名区域) | Variant() |
| [`SqlRangeQuery`](#sqlrangequery) | `(sql, rng, [tableAlias], [outOk])` | 对 Range 直接查询 (无需保存工作簿) | Variant() |

| Recipe | Functions Used |
|--------|---------------|
| [工作表查询](#recipe-worksheet-query) | `UDF_SQL_QUERY` |
| [多表关联](#recipe-multi-table-join) | `UDF_SQL_JOIN` |
| [分组聚合](#recipe-group-aggregation) | `UDF_SQL_GROUPBY` |

### Recipe 4.1 -- 工作表查询

<a id="recipe-worksheet-query"></a>

**场景**: 从学生成绩表中查询成绩大于 80 分的学生，并按成绩降序排列。

**输入** (工作表名: `Sheet1`)

|   | A | B | C |
|---|---|---|---|
| 1 | 学生 | 科目 | 成绩 |
| 2 | 张三 | 数学 | 85 |
| 3 | 李四 | 英语 | 92 |
| 4 | 王五 | 数学 | 78 |
| 5 | 赵六 | 语文 | 88 |

`=UDF_SQL_QUERY("SELECT * FROM [Sheet1$] WHERE 成绩 > 80 ORDER BY 成绩 DESC")` →

**输出**
|   | A | B | C |
|---|---|---|---|
| 1 | 学生 | 科目 | 成绩 |
| 2 | 李四 | 英语 | 92 |
| 3 | 赵六 | 语文 | 88 |
| 4 | 张三 | 数学 | 85 |

#### SqlExecute

执行任意 SQL 语句，返回 2D Variant 数组 (含列名标题行)。**仅 VBA** (通过 outOk 参数指示成功/失败)。

**VBA Usage**
```vb
SqlExecute(sql, [filePath], [includeHeader], outOk) As Variant()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| sql | String | SQL 语句 |
| filePath | String | 可选. 工作簿路径，默认当前工作簿 |
| includeHeader | Boolean | 可选. 是否包含列名标题行，默认 True |
| outOk | Boolean | 输出. 执行成功返回 True |

```vb
Dim ok As Boolean
result = SqlExecute("SELECT * FROM [Sheet1$] WHERE Age > 30", , , ok)
If ok Then Range("A1").Resize(UBound(result,1), UBound(result,2)).Value = result
```

**UDF Usage**
```
=UDF_SQL_QUERY(sql, [filePath])
```

| 参数 | 类型 | 说明 |
|------|------|------|
| sql | String | 完整的 SQL 语句 |
| filePath | String | 可选. 工作簿路径，默认当前工作簿 |

#### SqlQuery

SELECT 查询 (分句传入，自动拼接和转义表名)。

**VBA Usage**
```vb
SqlQuery(selectClause, fromClause, [whereClause], [orderByClause], [filePath], outOk) As Variant()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| selectClause | String | SELECT 子句 (如 "*" 或 "Name, Age") |
| fromClause | String | 表名 (自动添加 [$] 和方括号转义) |
| whereClause | String | 可选. WHERE 条件 |
| orderByClause | String | 可选. ORDER BY 子句 |
| filePath | String | 可选. 工作簿路径 |
| outOk | Boolean | 输出 |

```vb
result = SqlQuery("*", "Sheet1", "Age > 30", "Name ASC", , ok)
```

#### CloseSqlCache

释放模块级缓存的 ADODB 连接。建议在 Workbook_BeforeClose 中调用。

**VBA Usage**
```vb
CloseSqlCache
```

#### SqlGetConnection

获取 ADODB 数据库连接 (模块级 Static 缓存，同一文件复用连接)。优先 ACE (xlsx/xlsm)，回退 Jet (xls)。**仅 VBA**。

**VBA Usage**
```vb
SqlGetConnection([filePath], outOk) As Object
```

#### SqlRangeQuery

对 Range 直接查询，无需保存工作簿。使用 ADODB Recordset 在内存中操作。支持 SELECT * 和 WHERE 筛选。

**VBA Usage**
```vb
SqlRangeQuery(sql, rng, [tableAlias], outOk) As Variant()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| sql | String | SQL 语句 (SELECT * FROM ... WHERE ...) |
| rng | Range | 数据区域 (首行为列名) |
| tableAlias | String | 可选. 表别名，默认 "data" |
| outOk | Boolean | 输出 |

```vb
result = SqlRangeQuery("SELECT * FROM data WHERE Score > 90", Range("A1:C100"), "data", ok)
```

### Recipe 4.2 -- 多表关联

<a id="recipe-multi-table-join"></a>

**场景**: 销售表与区域表关联，查询各销售记录的所属区域名称。

**输入** (工作表: `Sales`)
|   | A | B |
|---|---|---|
| 1 | Product | RegionID |
| 2 | 苹果 | 1 |
| 3 | 香蕉 | 2 |

(工作表: `Regions`)
|   | A | B |
|---|---|---|
| 1 | ID | RegionName |
| 2 | 1 | 北京 |
| 3 | 2 | 上海 |

`=UDF_SQL_JOIN("Sales", "Regions", "t1.RegionID = t2.ID", "INNER", "t1.Product, t2.RegionName")` →

**输出**
|   | A | B |
|---|---|---|
| 1 | Product | RegionName |
| 2 | 苹果 | 北京 |
| 3 | 香蕉 | 上海 |

#### SqlJoin

两表 JOIN 查询。

**VBA Usage**
```vb
SqlJoin(table1, table2, joinOn, [joinType], [selectCols], [filePath], outOk) As Variant()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| table1 | String | 左表名 |
| table2 | String | 右表名 |
| joinOn | String | JOIN 条件 (如 "t1.ID = t2.ID") |
| joinType | String | 可选. "INNER" (默认), "LEFT", "RIGHT" |
| selectCols | String | 可选. 输出列，默认 "*" |
| filePath | String | 可选. 工作簿路径 |
| outOk | Boolean | 输出 |

```vb
result = SqlJoin("Sales", "Regions", "t1.RegionID = t2.ID", "LEFT", , , ok)
```

**UDF Usage**
```
=UDF_SQL_JOIN(table1, table2, joinOn, [joinType], [selectCols], [filePath])
```

### Recipe 4.3 -- 分组聚合

<a id="recipe-group-aggregation"></a>

**场景**: 按地区分组统计销售额的总和和记录数。

**输入** (工作表: `Sales`)
|   | A | B |
|---|---|---|
| 1 | Region | Amount |
| 2 | 北京 | 1000 |
| 3 | 上海 | 800 |
| 4 | 北京 | 1200 |
| 5 | 广州 | 600 |

`=UDF_SQL_GROUPBY("Sales", "Region", "SUM(Amount) AS Total, COUNT(*) AS Cnt")` →

**输出**
|   | A | B | C |
|---|---|---|---|
| 1 | Region | Total | Cnt |
| 2 | 北京 | 2200 | 2 |
| 3 | 上海 | 800 | 1 |
| 4 | 广州 | 600 | 1 |

#### SqlGroupBy

分组聚合查询。

**VBA Usage**
```vb
SqlGroupBy(tableName, groupCols, aggExprs, [whereClause], [filePath], outOk) As Variant()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| tableName | String | 表名 |
| groupCols | String | 分组列 (逗号分隔) |
| aggExprs | String | 聚合表达式 (如 "SUM(Amount) AS Total") |
| whereClause | String | 可选. WHERE 筛选条件 |
| filePath | String | 可选. 工作簿路径 |
| outOk | Boolean | 输出 |

```vb
result = SqlGroupBy("Sales", "Region", "SUM(Amount) AS Total, AVG(Amount) AS Avg", , , ok)
```

**UDF Usage**
```
=UDF_SQL_GROUPBY(tableName, groupCols, aggExprs, [whereClause], [filePath])
```

#### SqlListSheets

列出工作簿中所有的工作表名 (含 $ 后缀)。

**VBA Usage**
```vb
SqlListSheets([filePath], outOk) As Variant()
```

**UDF Usage**
```
=UDF_SQL_LIST_SHEETS([filePath])
```

#### SqlListColumns

列出指定表的列名及其序号 (ORDINAL_POSITION)。

**VBA Usage**
```vb
SqlListColumns(tableName, [filePath], outOk) As Variant()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| tableName | String | 表名 (自动添加 $ 后缀和方括号) |
| filePath | String | 可选. 工作簿路径 |

**UDF Usage**
```
=UDF_SQL_LIST_COLUMNS(tableName, [filePath])
```

#### SqlListTables

列出工作簿中所有可用数据源 (包括工作表和工作簿级命名区域)。

**VBA Usage**
```vb
SqlListTables([filePath], outOk) As Variant()
```

**UDF Usage**
```
=UDF_SQL_LIST_TABLES([filePath])
```

---

> **环境要求**: SqlUtils 依赖 ADODB 和 ACE/Jet OLEDB 提供程序。64 位 Office 需安装 Access Database Engine。若工作簿未保存，使用 SqlRangeQuery 可直接对 Range 查询。
>
> **⚠️ SQL 注入防护**: ACE OLEDB Excel ISAM 驱动不支持参数化查询。使用 `SqlEscapeString()` 转义来自单元格的用户输入（`'` → `''`），避免在 `WHERE`/`VALUES` 子句中直接拼接未转义的用户输入: `=UDF_SQL_QUERY("SELECT * FROM [Sheet1$] WHERE Name = '" & SqlEscapeString(A1) & "'")`。注意: Error 值单元格（#N/A、#VALUE! 等）会被静默转为空字符串（防 `CStr` 崩溃）——上游数据含 Error 时可能得到空结果集，请先清理数据。
>
> **命名约定**: 工作表名在 SQL 中需加 `$` 后缀并用方括号包裹，如 `[Sheet1$]`。列名含特殊字符时 SqlRangeQuery 会自动清理。
---

## Chapter 5: LinearUtils — 矩阵与线性代数计算

矩阵运算、分解、求解与工作表数组公式。所有输出矩阵均为 1-based (与 Excel Range 一致)。**模块**: `LinearUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`MatrixRows`](#matrixrows) | `(A)` | 矩阵行数 | Long |
| [`MatrixCols`](#matrixcols) | `(A)` | 矩阵列数 | Long |
| [`MatrixFrobeniusNorm`](#matrixnorm) | `(A)` | Frobenius 范数 (缩放过防溢出) | Double |
| [`RangeToMatrix`](#rangetomatrix) | `(rng)` | Range 转 Double(,) 1-based | Double(,) |
| [`MatrixToRange`](#matrixtorange) | `(mat, startCell)` | 矩阵写入工作表 (Sub) | — |
| [`SelectionToArray2D`](#rangetomatrix) | `(rng)` | 选中区域转 Variant 2D 数组 | Variant(,) |
| [`ArrayToRange`](#matrixtorange) | `(data, startCell, [asColumn])` | Variant 数组写入 Range (Sub) | — |
| [`MatrixTranspose`](#matrixtranspose) | `(A)` | 矩阵转置 | Double(,) |
| [`IdentityMatrix`](#identitymatrix) | `(n)` | 生成 n×n 单位矩阵 | Double(,) |
| [`MatrixCopy`](#matrixcopy) | `(A)` | 深拷贝矩阵 | Double(,) |
| [`MatrixGetColumn`](#matrixrows) | `(A, k)` | 提取第 k 列 (1-based) | Double() |
| [`MatrixSetColumn`](#matrixrows) | `(A, k, col)` | 设置第 k 列 (Sub) | — |
| [`MatrixMultiply`](#matrixmultiply) | `(A, B, [blockSize])` | 分块矩阵乘法 (缓存友好) | Double(,) |
| [`MatrixMultiplyNaive`](#matrixmultiply) | `(A, B)` | 朴素三重循环乘法 | Double(,) |
| [`MatrixScale`](#matrixscale) | `(A, scalar)` | 标量乘法 | Double(,) |
| [`MatrixAdd`](#matrixadd) | `(A, B)` | 矩阵加法 | Double(,) |
| [`MatrixSubtract`](#matrixsubtract) | `(A, B)` | 矩阵减法 A-B | Double(,) |
| [`MatrixHadamard`](#matrixhadamard) | `(A, B)` | 逐元素乘积 | Double(,) |
| [`MatrixPower`](#matrixpower) | `(A, n)` | 方阵 n 次幂 (二分法, n>=0) | Double(,) |
| [`MatrixNorm`](#matrixnorm) | `(A, [normType])` | 矩阵范数 ("1"/"inf"/"fro") | Double |
| [`MatrixTrace`](#matrixtrace) | `(A)` | 迹 (对角线之和) | Double |
| [`SVD`](#svd) | `(A, U, S, Vt, [tol])` | 奇异值分解 One-Sided Jacobi (Sub) | ByRef |
| [`PseudoInverse`](#pseudoinverse) | `(A, [tolerance])` | Moore-Penrose 伪逆 | Double(,) |
| [`MatrixRank_Array`](#matrixrank_array) | `(A, [tolerance])` | 基于 SVD 的矩阵秩 | Long |
| [`EigenSymmetric`](#eigensymmetric) | `(A, V, D, [tol])` | 对称矩阵特征分解 Jacobi (Sub) | ByRef |
| [`QRDecomposition`](#qrdecomposition) | `(A, Q, R, [economy])` | QR 分解 Householder (Sub) | ByRef |
| [`LUDecomposition`](#ludecomposition) | `(A, L, U, P, [swapCount])` | LU 分解 Doolittle 部分主元 (Sub) | ByRef |
| [`CholeskyDecomposition`](#choleskydecomposition) | `(A, L)` | Cholesky 分解 A=LL^T (Sub) | ByRef |
| [`MatrixDeterminant`](#matrixdeterminant) | `(A)` | 行列式 (通过 LU 分解) | Double |
| [`SolveLinearSystem`](#solvelinearsystem) | `(A, b, [tolerance])` | 解 Ax=b (SVD 伪逆) | Double() |
| [`MatrixConditionNumber`](#matrixconditionnumber) | `(A, [tol], [maxSweeps])` | 条件数 max(S)/min(S) | Double |
| [`VectorDot`](#vectordot) | `(a, b)` | 向量点积 (Kahan 补偿) | Double |
| [`VectorNorm`](#vectornorm) | `(v, [normType])` | 向量范数 ("2"/"1"/"inf") | Double |
| [`VectorCross`](#vectorcross) | `(a, b)` | 三维向量叉积 | Double(3) |
| [`PolyFit`](#polyfit) | `(rngX, rngY, [degree])` | 最小二乘多项式拟合 (QR 法) | Variant(,1) |

| Recipe | Functions Used |
|--------|---------------|
| [SVD 奇异值分解](#recipe-svd) | `UDF_LINALG_SVD_SVALS`, `UDF_LINALG_SVD_U`, `UDF_LINALG_SVD_VT` |
| [线性方程组求解](#recipe-linear-system) | `UDF_LINALG_SOLVE` |
| [多项式拟合](#recipe-polyfit) | `UDF_LINALG_POLYFIT` |

<a id="recipe-svd"></a>
### Recipe 5.1 — SVD 奇异值分解

**场景**: 对矩阵进行奇异值分解，获取奇异值、左右奇异向量，用于降维、伪逆计算或矩阵压缩。

**输入**
|   | A | B | C |
|---|---|---|---|
| 1 | 1 | 0 | 0 |
| 2 | 0 | 2 | 0 |
| 3 | 0 | 0 | 0 |

选中 3x1 区域输入数组公式 `=UDF_LINALG_SVD_SVALS(A1:C3)` → `{2; 1; 0}`

选中 3x3 区域输入 `=UDF_LINALG_SVD_U(A1:C3)` → 左奇异向量矩阵

`=UDF_LINALG_PINV(A1:C3, 1E-10)` → 伪逆 (3x3)

#### SVD

对任意 mxn 矩阵进行奇异值分解 A = U * S * V^T。使用 One-Sided Jacobi 算法，数值稳定。返回 U (mxn)、S (nxn 对角矩阵)、V^T (nxn)。

**VBA Usage**
```vb
SVD(A, U, S, Vt, [tol], [maxSweeps])
```

| 参数 | 类型 | 说明 |
|------|------|------|
| A | Double(,) | 输入矩阵 |
| U | Double(,) | 输出: 左奇异向量矩阵 (ByRef) |
| S | Double(,) | 输出: 奇异值对角矩阵 (ByRef) |
| Vt | Double(,) | 输出: 右奇异向量转置 (ByRef) |
| tol | Double | 可选. 收敛容差 |
| maxSweeps | Long | 可选. 最大扫描次数 |

```vb
Dim A() As Double, U() As Double, S() As Double, Vt() As Double
A = RangeToMatrix(Sheet1.Range("A1:C3"))
SVD A, U, S, Vt
MatrixToRange U, Sheet1.Range("E1")
' 当 m < n (宽矩阵) 时自动转置 A^T 计算再交换 U/Vt
```

> **SVD 提供什么**: A = U * S * V^T。**S** 包含奇异值（对角、降序）——每个分量的"重要性"。非零奇异值的个数等于矩阵有效秩。**U** 列是左奇异向量，**V^T** 行是右奇异向量。常见用途：降维（保留前 k 个奇异值）、伪逆 (`PseudoInverse`) 和类 PCA 分析。

**UDF Usage**
```
=UDF_LINALG_SVD_U(rng, [tol], [maxSweeps])       → U 矩阵 (左奇异向量)
=UDF_LINALG_SVD_S(rng, [tol], [maxSweeps])       → S 对角矩阵
=UDF_LINALG_SVD_VT(rng, [tol], [maxSweeps])      → V^T 矩阵 (右奇异向量转置)
=UDF_LINALG_SVD_SVALS(rng, [tol], [maxSweeps])   → 奇异值列向量
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rng | Range | 数值矩阵区域 |
| tol | Double | 可选. 收敛容差 |
| maxSweeps | Long | 可选. 最大扫描次数 |

#### PseudoInverse

基于 SVD 的 Moore-Penrose 伪逆。奇异值小于容差的被截断为零。负容差自动使用 max(S) * eps * max(m,n)。

**VBA Usage**
```vb
PseudoInverse(A, [tolerance]) As Double()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| A | Double(,) | 输入矩阵 |
| tolerance | Double | 可选. 奇异值截断容差 (-1=自动) |

```vb
Ap = PseudoInverse(A)
```

**UDF Usage**
```
=UDF_LINALG_PINV(rng, [tolerance])
```

#### EigenSymmetric

对称矩阵特征分解 A = V * D * V^T。使用 Jacobi 旋转算法。自动检查对称性，特征值降序排列。

**VBA Usage**
```vb
EigenSymmetric(A, V, D, [tol], [maxSweeps])
```

| 参数 | 类型 | 说明 |
|------|------|------|
| A | Double(,) | 对称方阵 |
| V | Double(,) | 输出: 特征向量矩阵 (ByRef) |
| D | Double(,) | 输出: 特征值对角矩阵 (ByRef) |

**UDF Usage**
```
=UDF_LINALG_EIGVAL(rng, [tol], [maxSweeps])   ' 特征值
=UDF_LINALG_EIGVEC(rng, [tol], [maxSweeps])   ' 特征向量
```

<a id="recipe-linear-system"></a>
### Recipe 5.2 — 线性方程组求解

**场景**: 求解线性方程组 Ax = b，支持方阵精确解、超定最小二乘解和欠定最小范数解。

**输入**
|   | A | B | C | D |
|---|---|---|---|---|
| 1 | 2 | 1 | -1 | 8 |
| 2 | -3 | -1 | 2 | -11 |
| 3 | -2 | 1 | 2 | -3 |

`=UDF_LINALG_SOLVE(A1:C3, D1:D3)` → `{2; 3; -1}` (解 x=2, y=3, z=-1)

#### SolveLinearSystem

使用 SVD 伪逆求解 Ax = b。方阵非奇异时得到精确解；超定 (m>n) 时得到最小二乘解；欠定 (m<n) 时得到最小范数解。

**VBA Usage**
```vb
SolveLinearSystem(A, b, [tolerance]) As Double()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| A | Double(,) | mxn 系数矩阵 |
| b | Double() | 右端向量 (长度 m) |
| tolerance | Double | 可选. SVD 容差 |

```vb
x = SolveLinearSystem(A, b)
' x(1)=2, x(2)=3, x(3)=-1
```

**UDF Usage**
```
=UDF_LINALG_SOLVE(rngA, rngB)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rngA | Range | 系数矩阵 A |
| rngB | Range | 右端向量 b (单列/单行) |

#### MatrixDeterminant

通过 LU 分解计算方阵行列式。符号由置换次数决定。

**VBA Usage**
```vb
MatrixDeterminant(A) As Double
```

```vb
det = MatrixDeterminant(A)
```

**UDF Usage**
```
=UDF_LINALG_DET(rng)
```

#### MatrixConditionNumber

条件数 = 最大奇异值 / 最小奇异值。值大表示矩阵病态。奇异时返回 MAX_DOUBLE。

**VBA Usage**
```vb
MatrixConditionNumber(A, [tol], [maxSweeps]) As Double
```

<a id="recipe-polyfit"></a>
### Recipe 5.3 — 多项式拟合

**场景**: 给定 x、y 数据点，用最小二乘法拟合多项式，返回系数 (高次到低次)。

**输入**
|   | A | B |
|---|---|---|
| 1 | 1 | 2.1 |
| 2 | 2 | 3.9 |
| 3 | 3 | 6.2 |
| 4 | 4 | 7.8 |
| 5 | 5 | 10.1 |

`=UDF_LINALG_POLYFIT(A1:A5, B1:B5, 1)` → `{2.01; 0.02}` (y = 2.01x + 0.02)

`=UDF_LINALG_POLYFIT(A1:A5, B1:B5, 2)` → `{0.01; 1.95; 0.13}` (y = 0.01x^2 + 1.95x + 0.13)

#### PolyFit

使用 QR 分解求解最小二乘问题，避免正规方程平方条件数的问题。degree=1 返回 [a,b] 对应 y=ax+b；degree=2 返回 [a,b,c] 对应 y=ax^2+bx+c。

**VBA Usage**
```vb
PolyFit(rngX, rngY, [degree]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rngX | Range/Array | 自变量数据 (1D) |
| rngY | Range/Array | 因变量数据 (1D, 与 X 等长) |
| degree | Long | 可选. 多项式次数 (默认 1) |

```vb
result = PolyFit(Range("A1:A5"), Range("B1:B5"), 2)
' result 为 2D 列向量, 高次到低次
```

**UDF Usage**
```
=UDF_LINALG_POLYFIT(rngX, rngY, [degree])
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rngX | Range | X 数据区域 |
| rngY | Range | Y 数据区域 (与 X 等行数) |
| degree | Long | 可选. 多项式次数 (默认 1) |

<a id="matrixrows"></a>
<a id="matrixcols"></a>
#### MatrixRows / MatrixCols

矩阵尺寸查询。

```vb
m = MatrixRows(A)   ' 行数
n = MatrixCols(A)   ' 列数
```

<a id="rangetomatrix"></a>
<a id="matrixtorange"></a>
#### RangeToMatrix / MatrixToRange

Range 与 Double 矩阵互转。RangeToMatrix 要求所有单元格为数值。

```vb
A = RangeToMatrix(Sheet1.Range("A1:C3"))    ' Range -> Double(,)
MatrixToRange U, Sheet1.Range("E1")          ' Double(,) -> Range
```

<a id="matrixtranspose"></a>
<a id="identitymatrix"></a>
<a id="matrixcopy"></a>
#### MatrixTranspose / IdentityMatrix / MatrixCopy

基本矩阵操作。

```vb
At = MatrixTranspose(A)        ' A^T
I = IdentityMatrix(3)          ' 3x3 单位矩阵
Acopy = MatrixCopy(A)          ' 深拷贝
```

<a id="matrixadd"></a>
<a id="matrixsubtract"></a>
<a id="matrixmultiply"></a>
<a id="matrixhadamard"></a>
#### MatrixAdd / MatrixSubtract / MatrixMultiply / MatrixHadamard

矩阵四则运算。MatrixMultiply 使用分块算法优化缓存。维度不匹配时抛出错误。

```vb
C = MatrixAdd(A, B)             ' A + B
D = MatrixSubtract(A, B)        ' A - B
E = MatrixMultiply(A, B, 64)    ' A * B (分块)
F = MatrixHadamard(A, B)        ' A (Hadamard) B 逐元素乘积
```

#### QRDecomposition

基于 Householder 反射的 QR 分解 A = Q*R。Sub 过程通过 ByRef 返回 Q 和 R。

**VBA Usage**
```vb
QRDecomposition A, Q, R, True      ' 经济模式: Q(mxk), R(kxk)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| A | Double(,) | 输入矩阵 |
| Q | Double(,) | 输出: 正交矩阵 (ByRef) |
| R | Double(,) | 输出: 上三角矩阵 (ByRef) |
| economy | Boolean | 可选. 经济模式 (默认 True) |

**UDF Usage**
```
=UDF_LINALG_QR_Q(rng, [economy])   → Q 矩阵
=UDF_LINALG_QR_R(rng, [economy])   → R 矩阵
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rng | Range | 数值矩阵区域 |
| economy | Boolean | 可选. 经济模式 (默认 False) |

#### QRDecompositionPiv

带列主元的 QR 分解 (Businger-Golub) A*P = Q*R。通过列置换提高数值稳定性，适用于秩亏检测、最小二乘和子集选择。

**VBA Usage**
```vb
QRDecompositionPiv A, Q, R, perm, True   ' 经济模式 + 列主元
```

| 参数 | 类型 | 说明 |
|------|------|------|
| A | Double(,) | 输入矩阵 |
| Q | Double(,) | 输出: 正交矩阵 (ByRef) |
| R | Double(,) | 输出: 上三角矩阵 (ByRef) |
| perm | Long() | 输出: 列置换向量 (ByRef), perm(k)=原始列索引 |
| economy | Boolean | 可选. 经济模式 (默认 False) |

#### CholeskyDecomposition

对称正定矩阵的 Cholesky 分解 A = L*L^T。Sub 过程通过 ByRef 返回 L (下三角)。

**VBA Usage**
```vb
CholeskyDecomposition A, L         ' A = L*L^T (下三角)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| A | Double(,) | 输入矩阵 (必须对称正定) |
| L | Double(,) | 输出: 下三角矩阵 (ByRef) |

**UDF Usage**
```
=UDF_LINALG_CHOLESKY(rng)          → L 矩阵 (下三角)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| rng | Range | 对称正定矩阵区域 |

#### LUDecomposition

Doolittle 部分主元法 LU 分解 P*A = L*U。Sub 过程通过 ByRef 返回 L, U, P 和行交换次数。

**VBA Usage**
```vb
LUDecomposition A, L, U, P, swaps  ' P*A = L*U
```

<a id="vectordot"></a>
<a id="vectornorm"></a>
<a id="vectorcross"></a>
#### VectorDot / VectorNorm / VectorCross

向量运算。点积使用 Kahan 补偿求和；叉积仅支持三维向量。

```vb
dot = VectorDot(v1, v2)            ' 点积 (Kahan 补偿)
nrm = VectorNorm(v, "2")           ' L2 范数 (缩放过)
cross = VectorCross(v1, v2)        ' 叉积 -> {x, y, z}
```

#### MatrixRank_Array

基于 SVD 奇异值容差的矩阵秩。

```vb
rank = MatrixRank_Array(A, 1E-10)
```

**UDF Usage**
```
=UDF_LINALG_RANK(rng, [tolerance])
```

<a id="matrixtrace"></a>
<a id="matrixnorm"></a>
<a id="matrixscale"></a>
<a id="matrixpower"></a>
#### MatrixTrace / MatrixNorm / MatrixScale / MatrixPower

矩阵标量运算。MatrixNorm 支持 "1" (列和范数)、"inf" (行和范数)、"fro" (Frobenius)。MatrixPower 使用二分幂运算。

```vb
tr = MatrixTrace(A)                ' 迹
n1 = MatrixNorm(A, "1")            ' 1-范数
B = MatrixScale(A, 3.5)            ' 标量乘法
Ap = MatrixPower(A, 3)             ' A^3 (二分法)
```

---

## Chapter 6: StatsUtils — 统计描述、推断与分布函数

集中趋势、离散度、形态、排名、关联分析、数据变换与假设检验。**模块**: `StatsUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`Mean`](#mean) | `(data, [colIndex])` | 算术平均值 (Kahan 补偿) | Variant (Double) |
| [`WeightedMean`](#mean) | `(values, weights)` | 加权算术平均 | Variant (Double) |
| [`Max`](#max) | `(data, [colIndex])` | 最大值 (VBA-only) | Variant (Double) |
| [`Median`](#median) | `(data, [colIndex])` | 中位数 | Variant (Double) |
| [`Min`](#min) | `(data, [colIndex])` | 最小值 (VBA-only) | Variant (Double) |
| [`MinMax`](#minmax) | `(data, [outMin], [outMax])` | 同时返回最小值/最大值 | Variant Array |
| [`Mode`](#mode) | `(data, [colIndex])` | 众数 (唯一时返回 #N/A) | Variant (Double) |
| [`GeometricMean`](#geometricmean) | `(data, [colIndex])` | 几何均值 (对数空间防溢出) | Variant (Double) |
| [`HarmonicMean`](#harmonicmean) | `(data, [colIndex])` | 调和均值 | Variant (Double) |
| [`TrimMean`](#trimmean) | `(data, [trimPct])` | 截尾均值 (剔除两端各 trimPct/2) | Variant (Double) |
| [`RootMeanSquare`](#rootmeansquare) | `(data, [colIndex])` | 均方根 RMS | Variant (Double) |
| [`MeanAbsDev`](#meanabsdev) | `(data, [colIndex])` | 平均绝对偏差 | Variant (Double) |
| [`StdDev`](#stddev) | `(data, [colIndex])` | 样本标准差 (n-1) | Variant (Double) |
| [`StdDevP`](#stddevp) | `(data, [colIndex])` | 总体标准差 (n) | Variant (Double) |
| [`Variance`](#variance) | `(data, [colIndex])` | 样本方差 (n-1, Kahan 补偿) | Variant (Double) |
| [`VarianceP`](#variancep) | `(data, [colIndex])` | 总体方差 (n) | Variant (Double) |
| [`Percentile`](#percentile) | `(data, k, [colIndex])` | 百分位数 (线性插值, 0<=k<=1) | Variant (Double) |
| [`IQR`](#iqr) | `(data, [colIndex])` | 四分位距 Q3-Q1 | Variant (Double) |
| [`Skewness`](#skewness) | `(data, [colIndex])` | 样本偏度 (Fisher-Pearson) | Variant (Double) |
| [`Kurtosis`](#kurtosis) | `(data, [colIndex])` | 超额峰度 | Variant (Double) |
| [`Rank`](#rank) | `(data, value, [ascending])` | 改良竞争排名 (ties 计为 <=) | Variant (Long) |
| [`RankEq`](#rankeq) | `(data, value, [ascending])` | 与 Excel RANK.EQ 一致 | Variant (Long) |
| [`RankAvg`](#rankavg) | `(data, value, [ascending])` | 与 Excel RANK.AVG 一致 | Variant (Double) |
| [`PercentRank`](#percentrank) | `(data, value, [ascending])` | 百分位排名 (0-1) | Variant (Double) |
| [`Covariance`](#covariance) | `(dataX, dataY)` | 样本协方差 (成对删除) | Variant (Double) |
| [`Correlation`](#correlation) | `(dataX, dataY)` | Pearson 相关系数 | Variant (Double) |
| [`RSquare`](#rsquare) | `(actual, predicted)` | R^2 判定系数 | Variant (Double) |
| [`CorrelationMatrix`](#correlationmatrix) | `(data, [hasHeader])` | 数值列相关系数矩阵 (含标签) | Variant(,) |
| [`ZScore`](#zscore) | `(data, [value])` | Z 标准化 (给出 value 返回单值) | Variant |
| [`Normalize`](#normalize) | `(data, [colIndex])` | Min-Max 归一化到 [0, 1] | Double() |
| [`LinInterp`](#lininterp) | `(x, xs, ys)` | 线性插值 (超出范围返回边界) | Variant (Double) |
| [`Winsorize`](#winsorize) | `(data, [pct])` | 缩尾处理 (两端各 pct/2) | Double() |
| [`MovingAverage`](#movingaverage) | `(data, window)` | 简单移动平均 | Double() |
| [`Binning`](#binning) | `(data, nBins)` | 等宽分箱 **仅 VBA** | Dictionary |
| [`ZTest`](#ztest) | `(data, mu0, sigma)` | 单样本 Z 检验 (双侧 p 值) | Variant (Double) |
| [`TTest`](#ttest) | `(data1, data2, [testType])` | 双样本 t 检验 (1=配对/2=等方差/3=Welch) | Variant (Double) |
| [`StandardError`](#standarderror) | `(data, [colIndex])` | 均值标准误差 SE = s/sqrt(n) | Variant (Double) |
| [`ConfidenceInterval`](#confidenceinterval) | `(data, [alpha])` | t 分布置信区间 **仅 VBA** | Dictionary |
| [`GammaLn`](#gammaln) | `(x)` | 对数 Gamma 函数 (Lanczos) | Double |
| [`BetaReg`](#betareg) | `(x, a, b)` | 正则化不完全 Beta 函数 I_x(a,b) | Double |
| [`TDistCDF`](#tdistcdf) | `(tVal, df)` | t 分布上尾概率 P(T>|t|) | Double |
| [`TDist2T`](#tdist2t) | `(tStat, df)` | t 分布双尾 P 值 2*P(T>|t|) | Double |
| [`FDistRT`](#fdistrt) | `(fStat, df1, df2)` | F 分布右尾概率 P(F>f) | Double |
| [`TInv2T`](#tinv2t) | `(alpha, df)` | t 分布双侧临界值 (二分搜索) | Double |

| Recipe | Functions Used |
|--------|---------------|
| [描述统计](#recipe-descriptive-stats) | `UDF_STAT_MEAN`, `UDF_STAT_STDEV`, `UDF_STAT_MEDIAN`, `UDF_STAT_IQR` |
| [假设检验](#recipe-hypothesis-test) | `UDF_STAT_ZTEST`, `UDF_STAT_TTEST` |
| [相关性分析](#recipe-correlation) | `UDF_STAT_CORREL`, `UDF_STAT_COV` |

#### Min

#### Max

返回数组中的最小值和最大值。无数值时返回 Empty。**仅 VBA**。

**VBA Usage**

```vb
Min(data, [colIndex]) As Variant
Max(data, [colIndex]) As Variant
```

| Parameter | Type | Description |
|-----------|------|-------------|
| data | Variant | 1D 或 2D 数组 |
| colIndex | Long | 可选。2D 数组的列索引(1-based)，默认第 1 列 |

```vb
Min(Array(3, 1, 5, 2))   ' → 1
Max(Array(3, 1, 5, 2))   ' → 5
```

<a id="recipe-descriptive-stats"></a>
### Recipe 6.1 — 描述统计

**场景**: 对一组实验数据计算均值、标准差、中位数、四分位距，快速了解数据分布特征。

**输入**
|   | A |
|---|---|
| 1 | 12.5 |
| 2 | 13.1 |
| 3 | 11.8 |
| 4 | 14.2 |
| 5 | 12.9 |
| 6 | 45.0 |
| 7 | 13.3 |

`=UDF_STAT_MEAN(A1:A7)` → `17.54`

`=UDF_STAT_MEDIAN(A1:A7)` → `13.1`

`=UDF_STAT_STDEV(A1:A7)` → `12.13` (样本标准差，受异常值 45.0 影响)

`=UDF_STAT_IQR(A1:A7)` → `1.05` (四分位距，对异常值稳健)

`=UDF_STAT_SKEW(A1:A7)` → `2.14` (严重右偏)

`=UDF_STAT_PERCENTILE(A1:A7, 0.9)` → `45.0` (第 90 百分位数)

> **如何解读这些结果**: 均值 (17.54) 被异常值 45.0 拉高，高估了典型值。中位数 (13.1) 不受异常值影响，此处更能代表数据中心。均值与中位数差距大是数据偏态的警示信号。IQR (1.05) 说明中间 50% 的数据跨度仅 1.05 单位——除异常值外数据非常集中。偏度 >1 或 <-1 表明严重偏态，接近 0 为对称分布。处理样本数据用 StdDev (n-1)，处理全部总体数据用 StdDevP (n)。

<a id="mean"></a>
<a id="median"></a>
<a id="mode"></a>
<a id="minmax"></a>
#### Mean / Median / Mode / MinMax

集中趋势统计量。Mean 使用 Kahan 补偿求和。Mode 无重复值时返回 #N/A。MinMax 同时返回最小值和最大值，适合数组公式。

**VBA Usage**
```vb
Mean(data, [colIndex]) As Variant
Median(data, [colIndex]) As Variant
Mode(data, [colIndex]) As Variant
MinMax(data, [outMin], [outMax], [colIndex]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| data | Range/Array | 数值数据 |
| colIndex | Long | 可选. 2D 数组时的列号 (默认 1) |

```vb
avg = Mean(Range("A1:A100"))
med = Median(Array(1, 2, 3, 4, 100))  ' -> 3
mm = MinMax(Range("A1:A100"))          ' -> Array(min, max)
```

**UDF Usage** (selected)
```
=UDF_STAT_MEAN(data, [colIndex])
=UDF_STAT_MEDIAN(data, [colIndex])
=UDF_STAT_MODE(data, [colIndex])
=UDF_STAT_MINMAX(data, [colIndex])
```

<a id="stddev"></a>
<a id="stddevp"></a>
<a id="variance"></a>
<a id="variancep"></a>
#### StdDev / StdDevP / Variance / VarianceP

离散度统计量。样本版本除以 n-1，总体版本除以 n。均使用 Kahan 补偿求和。

**VBA Usage**
```vb
StdDev(data, [colIndex]) As Variant      ' 样本标准差
StdDevP(data, [colIndex]) As Variant     ' 总体标准差
Variance(data, [colIndex]) As Variant    ' 样本方差
VarianceP(data, [colIndex]) As Variant   ' 总体方差
```

```vb
s = StdDev(Range("A1:A100"))     ' 样本标准差
s2 = Variance(Range("A1:A100"))  ' 样本方差
```

**UDF Usage**
```
=UDF_STAT_STDEV(data, [colIndex])
=UDF_STAT_STDEVP(data, [colIndex])
=UDF_STAT_VAR(data, [colIndex])
=UDF_STAT_VARP(data, [colIndex])
```

<a id="geometricmean"></a>
<a id="harmonicmean"></a>
<a id="trimmean"></a>
<a id="rootmeansquare"></a>
<a id="meanabsdev"></a>
#### GeometricMean / HarmonicMean / TrimMean / RootMeanSquare / MeanAbsDev

其他均值指标。GeometricMean 适合增长率场景；HarmonicMean 适合速度/比率场景；TrimMean 剔除极端值；RMS 用于信号强度；MAD 对异常值更稳健。

**VBA Usage**
```vb
GeometricMean(data, [colIndex]) As Variant
HarmonicMean(data, [colIndex]) As Variant
TrimMean(data, [trimPct], [colIndex]) As Variant
RootMeanSquare(data, [colIndex]) As Variant
MeanAbsDev(data, [colIndex]) As Variant
```

```vb
g = GeometricMean(Array(1.05, 1.08, 1.03))  ' -> ~1.053
h = HarmonicMean(Array(2#, 3#))              ' -> 2.4
t = TrimMean(Range("A1:A100"), 0.1)          ' 剔除两端各 5%
```

**UDF Usage**
```
=UDF_STAT_GEOMEAN(data, [colIndex])
=UDF_STAT_HARMEAN(data, [colIndex])
=UDF_STAT_TRIMEAN(data, [trimPct], [colIndex])
=UDF_STAT_RMS(data, [colIndex])
=UDF_STAT_MAD(data, [colIndex])
```

<a id="percentile"></a>
<a id="iqr"></a>
<a id="skewness"></a>
<a id="kurtosis"></a>
#### Percentile / IQR / Skewness / Kurtosis

分布形态与分位数。Percentile 使用线性插值法；Skewness >0 右偏，<0 左偏；Kurtosis >0 尖峰厚尾。

**VBA Usage**
```vb
Percentile(data, k, [colIndex]) As Variant   ' k in [0,1]
IQR(data, [colIndex]) As Variant
Skewness(data, [colIndex]) As Variant
Kurtosis(data, [colIndex]) As Variant
```

```vb
p90 = Percentile(Range("A1:A100"), 0.9)
iqrVal = IQR(Range("A1:A100"))
skew = Skewness(Range("A1:A100"))
kurt = Kurtosis(Range("A1:A100"))
```

**UDF Usage**
```
=UDF_STAT_PERCENTILE(data, k, [colIndex])
=UDF_STAT_IQR(data, [colIndex])
=UDF_STAT_SKEW(data, [colIndex])
=UDF_STAT_KURTOSIS(data, [colIndex])
```

<a id="recipe-hypothesis-test"></a>
### Recipe 6.2 — 假设检验

**场景**: 已知总体标准差 sigma=5，检验样本均值是否显著偏离 mu0=100；或比较两组数据的均值差异。

**输入**
|   | A |
|---|---|
| 1 | 98 |
| 2 | 102 |
| 3 | 95 |
| 4 | 100 |
| 5 | 97 |
| 6 | 103 |
| 7 | 99 |

Z 检验: sigma=5, mu0=100

`=UDF_STAT_ZTEST(A1:A7, 100, 5)` → `0.596` (p > 0.05，不拒绝 H0)

配对 t 检验，比较 A 组与 B 组:

|   | A | B |
|---|---|---|
| 1 | 85 | 88 |
| 2 | 90 | 92 |
| 3 | 78 | 82 |
| 4 | 92 | 95 |
| 5 | 88 | 91 |

`=UDF_STAT_TTEST(A1:A5, B1:B5, 1)` → `0.002` (p < 0.01，显著差异)

> **如何解读 p 值**: p 值是在"无真实效应"假设下观察到当前数据（或更极端）的概率。p 越小 = 反对原假设的证据越强。
> - **p < 0.01**: 强证据——差异具有统计显著性。
> - **p < 0.05**: 中等证据——常用显著性阈值。
> - **p > 0.05**: 证据不足——不能拒绝原假设，观察到的差异可能来自随机波动。
>
> **选择哪种检验**: 已知总体标准差时用 `ZTest`（适用于大样本）。小样本或标准差未知时用 `TTest`。`testType=1` 配对数据（同组前后对比），`testType=2` 等方差两组，`testType=3` (Welch) 不等方差两组。

#### ZTest

单样本 Z 检验，双侧 p 值 = 2 * (1 - Phi(|z|))。使用 Abramowitz-Stegun 正态 CDF 近似。

**VBA Usage**
```vb
ZTest(data, mu0, sigma, [colIndex]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| data | Range/Array | 样本数据 |
| mu0 | Double | 原假设均值 |
| sigma | Double | 已知总体标准差 |
| colIndex | Long | 可选. 2D 数组列号 |

```vb
p = ZTest(Range("A1:A30"), 100#, 5#)
' p < 0.05 -> 拒绝 H0
```

**UDF Usage**
```
=UDF_STAT_ZTEST(data, mu0, sigma, [colIndex])
```

#### TTest

双样本 t 检验。testType: 1=配对，2=等方差 (默认)，3=Welch。p 值基于纯 VBA t 分布 CDF (通过不完全 Beta 函数)。

**VBA Usage**
```vb
TTest(data1, data2, [testType], [colIdx1], [colIdx2]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| data1 | Range/Array | 第一组数据 |
| data2 | Range/Array | 第二组数据 |
| testType | Long | 可选. 1=配对, 2=等方差 (默认), 3=Welch |

```vb
p = TTest(Range("A1:A20"), Range("B1:B20"), 2)
```

**UDF Usage**
```
=UDF_STAT_TTEST(data1, data2, [testType])
```

<a id="standarderror"></a>
<a id="confidenceinterval"></a>
#### StandardError / ConfidenceInterval

标准误差 SE = s / sqrt(n)。ConfidenceInterval 返回 t 分布置信区间字典 (仅 VBA)，含 lower/upper/mean/se 键。

**VBA Usage**
```vb
StandardError(data, [colIndex]) As Variant
ConfidenceInterval(data, [alpha], [colIndex]) As Object  ' 仅 VBA
```

```vb
se = StandardError(Range("A1:A30"))
Set ci = ConfidenceInterval(Range("A1:A30"), 0.05)
' ci("lower") -> 下限, ci("upper") -> 上限
```

**UDF Usage**
```
=UDF_STAT_SE(data, [colIndex])
```

<a id="recipe-correlation"></a>
### Recipe 6.3 — 相关性分析

**场景**: 分析两组数值数据之间的线性相关性强度和方向。

**输入**
|   | A | B |
|---|---|---|
| 1 | 1 | 2.1 |
| 2 | 2 | 4.0 |
| 3 | 3 | 6.3 |
| 4 | 4 | 7.9 |
| 5 | 5 | 10.2 |

`=UDF_STAT_CORREL(A1:A5, B1:B5)` → `0.999` (极强正相关)

`=UDF_STAT_COV(A1:A5, B1:B5)` → `4.03` (正协方差)

`=UDF_STAT_R2(A1:A5, B1:B5)` → `0.998` (拟合良好)

> **如何解读相关性**: Pearson 相关系数 `r` 范围 -1 到 +1。**r > 0.8** 或 **r < -0.8** 强相关；**0.5 < |r| < 0.8** 中等；**|r| < 0.3** 弱相关。R² (决定系数) 表示 Y 的方差中有多少比例被 X 解释——此处 R²=0.998 意味着 B 的变异中有 99.8% 可由 A 解释。协方差表明关系方向，但其大小依赖单位；用相关系数做标准化比较。

<a id="correlation"></a>
<a id="covariance"></a>
<a id="rsquare"></a>
#### Correlation / Covariance / RSquare

关联分析。Correlation 为 Pearson 相关系数 r；Covariance 为样本协方差；RSquare 为 R^2 = 1 - SSE/SST。均使用成对删除处理缺失值。

**VBA Usage**
```vb
Covariance(dataX, dataY, [colX], [colY]) As Variant
Correlation(dataX, dataY, [colX], [colY]) As Variant
RSquare(actual, predicted, [colActual], [colPredicted]) As Variant
```

| 参数 | 类型 | 说明 |
|------|------|------|
| dataX | Range/Array | 第一变量 |
| dataY | Range/Array | 第二变量 |
| colX/colY | Long | 可选. 2D 数组时的列号 |

```vb
r = Correlation(Range("A1:A100"), Range("B1:B100"))
covXY = Covariance(Range("A1:A100"), Range("B1:B100"))
r2 = RSquare(actual, predicted)
```

**UDF Usage**
```
=UDF_STAT_CORREL(dataX, dataY, [colX], [colY])
=UDF_STAT_COV(dataX, dataY, [colX], [colY])
=UDF_STAT_R2(actual, predicted, [colActual], [colPredicted])
```

#### CorrelationMatrix

计算多列数值数据之间的 Pearson 相关系数矩阵。使用成对完全观测 (pairwise complete)，Kahan 补偿求和。输出含行/列标签的 2D Variant 数组。

**VBA Usage**
```vb
CorrelationMatrix(data, [hasHeader]) As Variant()
```

```vb
Dim corr As Variant
corr = CorrelationMatrix(Range("A1").CurrentRegion)
' 输出: 含 "Correlation" 标签矩阵
```

**UDF Usage** (通过 RegressUtils 代理)
```
=UDF_REGRESS_CORREL(data)
```

<a id="rank"></a>
<a id="rankeq"></a>
<a id="rankavg"></a>
<a id="percentrank"></a>
#### Rank / RankEq / RankAvg / PercentRank

排名函数。Rank 为改良竞争排名；RankEq 等价 Excel RANK.EQ；RankAvg 等价 RANK.AVG (ties 取平均)；PercentRank 返回 (rank-1)/(n-1)。

**VBA Usage**
```vb
Rank(data, value, [ascending], [colIndex]) As Variant
RankEq(data, value, [ascending], [colIndex]) As Variant
RankAvg(data, value, [ascending], [colIndex]) As Variant
PercentRank(data, value, [ascending], [colIndex]) As Variant
```

**UDF Usage**
```
=UDF_STAT_RANK(data, value, [ascending])
=UDF_STAT_RANKEQ(data, value, [ascending])
=UDF_STAT_RANKAVG(data, value, [ascending])
=UDF_STAT_PERCENTRANK(data, value, [ascending])
```

<a id="zscore"></a>
<a id="normalize"></a>
<a id="lininterp"></a>
<a id="winsorize"></a>
<a id="movingaverage"></a>
<a id="binning"></a>
#### ZScore / Normalize / LinInterp / Winsorize / MovingAverage / Binning

数据变换工具。ZScore 传入单值返回单 Z 分数，否则返回全数组；Normalize 为 Min-Max 归一化；LinInterp 在已排序 xs、ys 间线性插值；Winsorize 压缩极端值；MovingAverage 为简单移动平均；Binning 为等宽分箱 (仅 VBA 返回 Dictionary)。

**VBA Usage**
```vb
ZScore(data, [value], [colIndex]) As Variant
Normalize(data, [colIndex]) As Variant
LinInterp(x, xs, ys) As Variant
Winsorize(data, [pct], [colIndex]) As Variant
MovingAverage(data, window, [colIndex]) As Variant
Binning(data, nBins, [colIndex]) As Object        ' 仅 VBA
```

```vb
z = ZScore(Range("A1:A100"), 75#)
norm = Normalize(Range("A1:A100"))
y = LinInterp(2.5, Array(1,2,3), Array(10,20,30))  ' -> 25
w = Winsorize(Range("A1:A100"), 0.05)
ma = MovingAverage(Range("A1:A100"), 5)
Set bins = Binning(Range("A1:A100"), 10)
' bins("__internal_edges__") -> 边界数组
```

**UDF Usage** (selected)
```
=UDF_STAT_ZSCORE(data, [value], [colIndex])
=UDF_STAT_NORMALIZE(data, [colIndex])
=UDF_STAT_LININTERP(x, xs, ys)
=UDF_STAT_WINSORIZE(data, [pct], [colIndex])
=UDF_STAT_MA(data, window, [colIndex])
=UDF_STAT_WMEAN(values, weights, [colIdxV], [colIdxW])
```

#### GammaLn

对数 Gamma 函数 ln(Γ(x))。基于 Lanczos 近似 (g=7, 系数 9 项)，精度 ~1E-12。纯 VBA 实现，无 Excel 依赖。

```vb
GammaLn(x) As Double
```

| 参数 | 类型 | 说明 |
|------|------|------|
| x | Double | 正值，x > 0 |

```vb
gl = GammaLn(1#)     ' → 0      (ln(Γ(1)) = ln(1))
gl = GammaLn(5#)     ' → 3.178  (ln(24))
gl = GammaLn(0.5)    ' → 0.572  (ln(√π))
```

#### BetaReg

正则化不完全 Beta 函数 I_x(a,b)。使用 Modified Lentz 连分式展开，精度 ~1E-12。纯 VBA 实现。

```vb
BetaReg(x, a, b) As Double
```

| 参数 | 类型 | 说明 |
|------|------|------|
| x | Double | 积分上限，0 ≤ x ≤ 1 |
| a | Double | Beta 形状参数 1，a > 0 |
| b | Double | Beta 形状参数 2，b > 0 |

```vb
betaI = BetaReg(0.5, 2#, 3#)    ' → ~0.6875
betaI = BetaReg(1#, 1#, 1#)     ' → 1.0    (均匀分布)
```

#### TDistCDF

t 分布上尾概率 P(T > |t|)。通过 BetaReg 实现。纯 VBA，替代 Excel TDIST。

```vb
TDistCDF(tVal, df) As Double
```

| 参数 | 类型 | 说明 |
|------|------|------|
| tVal | Double | t 统计量，tVal ≥ 0 |
| df | Double | 自由度，df > 0 |

```vb
p = TDistCDF(2.5, 15)   ' P(T > 2.5), df=15
p = TDistCDF(1.96, 100) ' P(T > 1.96), df=100
```

#### TDist2T

t 分布双尾 P 值 2 × P(T > |t|)。用于双侧 t 检验。

```vb
TDist2T(tStat, df) As Double
```

| 参数 | 类型 | 说明 |
|------|------|------|
| tStat | Double | t 统计量 |
| df | Double | 自由度，df > 0 |

```vb
p2 = TDist2T(2.5, 15)   ' 双侧 P 值, df=15
p2 = TDist2T(3.0, 20)   ' 双侧 P 值, df=20
```

#### FDistRT

F 分布右尾概率 P(F > f)。通过 BetaReg 实现。纯 VBA，替代 Excel FDIST。

```vb
FDistRT(fStat, df1, df2) As Double
```

| 参数 | 类型 | 说明 |
|------|------|------|
| fStat | Double | F 统计量，fStat ≥ 0 |
| df1 | Double | 分子自由度，df1 > 0 |
| df2 | Double | 分母自由度，df2 > 0 |

```vb
pF = FDistRT(3.2, 3, 20)    ' P(F > 3.2), df=(3,20)
pF = FDistRT(5.0, 2, 30)    ' P(F > 5.0), df=(2,30)
```

#### TInv2T

t 分布双侧临界值。给定显著性水平 α，返回 t 临界值使得 P(|T| > t) = α。使用二分搜索，精度 ~1E-12。

```vb
TInv2T(alpha, df) As Double
```

| 参数 | 类型 | 说明 |
|------|------|------|
| alpha | Double | 显著性水平，0 < α < 1 |
| df | Double | 自由度，df > 0 |

```vb
tCrit = TInv2T(0.05, 20)    ' t 临界值 (α=0.05 双侧, df=20)
tCrit = TInv2T(0.01, 10)    ' t 临界值 (α=0.01 双侧, df=10)
```

---

## Chapter 7: RegressUtils — 回归建模、方差分析与因子优化

多元线性回归 (含分类/布尔因子编码)、单因素方差分析、交互效应检测、因子扫描与网格优化。**模块**: `RegressUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`FactorImportance`](#factorimportance) | `(data, factorCols, resultCol)` | 因子重要性排序 (标准化系数) | Variant(,) |
| [`InteractionEffects`](#interactioneffects) | `(data, factorCols, resultCol)` | 成对交互效应 F 检验 | Variant(,) |
| [`ANOVAOneWay`](#anovaoneway) | `(data, factorCol, resultCol)` | 单因素方差分析 **仅 VBA** | Dictionary |
| [`ANOVAOneWay_Fstat`](#anovaoneway) | `(data, factorCol, resultCol)` | 单因素方差分析 F 统计量 | Double |
| [`LinearModelFit`](#linearmodelfit) | `(data, factorCols, resultCol)` | 多元线性回归拟合 **仅 VBA** | Dictionary |
| [`LinearModelPredict`](#linearmodelpredict) | `(model, newData)` | 基于模型预测结果 | Double |
| [`FactorSweep`](#factorsweep) | `(model, factorIdx, from, to, steps)` | 单因子 What-If 扫描 | Variant(,) |
| [`OptimizeFactors`](#optimizefactors) | `(data, factorCols, resultCol, [goal], [topN])` | 网格搜索最优因子组合 | Variant(,) |

| Recipe | Functions Used |
|--------|---------------|
| [因子重要性分析](#recipe-factor-importance) | `UDF_REGRESS_IMPORTANCE` |
| [单因素方差分析](#recipe-one-way-anova) | `UDF_REGRESS_ANOVA` |
| [因子优化](#recipe-factor-optimization) | `UDF_REGRESS_OPTIMIZE` |

<a id="recipe-factor-importance"></a>
### Recipe 7.1 — 因子重要性分析

**场景**: 有多个数值因子影响产品质量指标，需要量化各因子的相对重要性，找出关键驱动因素。

**输入**
|   | A | B | C | D |
|---|---|---|---|---|
| 1 | 温度 | 压力 | 时间 | 产率 |
| 2 | 100 | 2.5 | 30 | 85.2 |
| 3 | 110 | 2.8 | 35 | 88.1 |
| 4 | 120 | 3.0 | 40 | 91.3 |
| 5 | 130 | 3.2 | 45 | 94.5 |
| 6 | 140 | 3.5 | 50 | 97.8 |

选中因子列名区域 (如 A1:C1) 作为 factorCols 参数：

`=UDF_REGRESS_IMPORTANCE(A1:D6, A1:C1, 4)` ->

| 排名 | 因子 | 标准化系数 | 原始系数 | 绝对重要性 |
|------|------|-----------|---------|-----------|
| 1 | 温度 | 0.852 | 0.321 | 0.852 |
| 2 | 压力 | 0.456 | 15.23 | 0.456 |
| 3 | 时间 | 0.213 | 0.089 | 0.213 |

#### FactorImportance

基于标准化回归系数对因子进行重要性排序。自动检测因子类型 (数值/布尔/分类)，分类因子自动编码为 k-1 个哑变量后聚合。输出含表头的 2D 数组：排名、因子名、标准化系数、原始系数、绝对重要性。

> **如何解读输出**:
> - **排名**: 1 = 最重要因子，按绝对重要性降序排列。
> - **标准化系数**: 系数缩放到 [-1, 1]，便于跨不同单位因子公平比较。绝对值越大 = 影响越大。
> - **原始系数**: 原始单位的回归系数。温度的 raw 系数 0.321 表示每升高 1 度产率增加 0.321 单位。
> - **绝对重要性**: 当前实现中与标准化系数相同。用此列排序因子重要性。
>
> 示例中：温度 (0.852) 的重要性约为压力 (0.456) 的 2 倍、时间 (0.213) 的 4 倍。

**VBA Usage**
```vb
FactorImportance(data, factorCols, resultCol, [hasHeader]) As Variant()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| data | Range/Variant(,) | 数据区域 (首行默认表头) |
| factorCols | Variant Array | 因子列索引数组 (1-based) |
| resultCol | Long | 结果列索引 (1-based) |
| hasHeader | Boolean | 可选. 是否含表头 (默认 True) |

```vb
Dim importance As Variant
importance = FactorImportance(dataRange, Array(1,2,3), 4)
' 输出含表头，标准化系数越大 -> 因子越重要
```

**UDF Usage**
```
=UDF_REGRESS_IMPORTANCE(data, factorCols, resultCol)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| data | Range | 完整数据区域 (含表头) |
| factorCols | Range | 因子列名区域 (如 A1:C1) |
| resultCol | Long | 结果列号 (1-based) |

<a id="recipe-one-way-anova"></a>
### Recipe 7.2 — 单因素方差分析

**场景**: 比较三个不同工艺参数设置下的产品强度是否有显著差异。

**输入**
|   | A | B |
|---|---|---|
| 1 | 工艺 | 强度 |
| 2 | A | 23 |
| 3 | A | 25 |
| 4 | A | 22 |
| 5 | B | 30 |
| 6 | B | 32 |
| 7 | B | 31 |
| 8 | C | 28 |
| 9 | C | 27 |
| 10 | C | 29 |

`=UDF_REGRESS_ANOVA(A1:B10, 1, 2)` -> `"F(2,6) = 25.500, p = 0.0012, eta^2 = 0.895"`

#### ANOVAOneWay

单因素方差分析，检验一个分类因子对数值结果是否有显著影响。使用两阶段组内平方和计算避免灾难性抵消。返回 Dictionary (仅 VBA) 含 SSB/SSW/SST/MSB/MSW/F/p_value/eta_sq/significant/summary。

**VBA Usage**
```vb
ANOVAOneWay(data, factorCol, resultCol, [hasHeader]) As Object  ' 仅 VBA
```

| 参数 | 类型 | 说明 |
|------|------|------|
| data | Range/Variant(,) | 数据区域 |
| factorCol | Long | 分类因子列号 (1-based) |
| resultCol | Long | 数值结果列号 (1-based) |
| hasHeader | Boolean | 可选. 默认 True |

`ANOVAOneWay_Fstat` 是便捷包装函数，仅返回 F 统计量 (Double)，参数相同。

```vb
Set anova = ANOVAOneWay(dataRange, 1, 2)
f = ANOVAOneWay_Fstat(dataRange, 1, 2)   ' → 8.42
' anova("summary") -> "F(2,6) = 25.500, p = 0.0012, eta^2 = 0.895"
' anova("F") -> 25.5
' anova("p_value") -> 0.0012
' anova("significant") -> True
```

> **如何解读输出**: Dictionary 包含：`SSB`组间/`SSW`组内/`SST`总平方和；`MSB`/`MSW`均方；`F`统计量（越大越显著）；`p_value`（**<0.05 组间差异显著**）；`eta_sq`效应量（>0.14 大，0.06~0.14 中，<0.01 小）；`significant`布尔值；`summary`可读摘要。上例 F(2,6)=25.5, p=0.0012 表明三种工艺差异极显著，eta²=0.895 表示 89.5% 的产品强度差异可由工艺选择解释。

**UDF Usage**
```
=UDF_REGRESS_ANOVA(data, factorCol, resultCol)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| data | Range | 数据区域 (含表头) |
| factorCol | Long | 分类因子列号 |
| resultCol | Long | 数值结果列号 |

<a id="linearmodelfit"></a>
<a id="linearmodelpredict"></a>
#### LinearModelFit / LinearModelPredict

多元线性回归拟合与预测。构建含截距、数值因子、布尔因子、分类哑变量的设计矩阵，使用 QR 分解求解系数，(X^T X)^(-1) 伪逆计算标准误差。返回 Dictionary 含 coefficients/coef_names/r_squared/adj_r_squared/se/t_stats/p_values/f_stat/f_pvalue/fitted_values/residuals/factor_map/formula。LinearModelPredict 基于模型预测新数据。

**VBA Usage**
```vb
LinearModelFit(data, factorCols, resultCol, [hasHeader]) As Object  ' 仅 VBA
LinearModelPredict(model, newData) As Double
```

| 参数 | 类型 | 说明 |
|------|------|------|
| data | Range/Variant(,) | 训练数据 |
| factorCols | Variant | 因子列索引数组 |
| resultCol | Long | 结果列索引 |
| newData | Range/Array | 预测用新因子值 (按 factorCols 顺序) |

```vb
Set model = LinearModelFit(dataRange, Array(1,2,3), 4)
' model("r_squared") -> R^2
' model("formula") -> "Y ~ 温度 + 压力 + 时间"
' model("coefficients") -> 系数数组

pred = LinearModelPredict(model, Array(120, 3.0, 40))
' -> 预测值
```

> **如何解读输出**: Dictionary 包含：`coefficients`回归系数（首值=截距）；`coef_names`因子名；`r_squared`（**>0.7 良好，>0.9 优秀**）；`adj_r_squared`（与 r_squared 差距大=多余因子）；`se`标准误；`t_stats`（**|t|>2 因子有意义**）；`p_values`（**<0.05 该因子显著**）；`f_stat`/`f_pvalue`整体 F 检验；`fitted_values`/`residuals`拟合值与残差；`factor_map`分类编码；`formula`公式字符串。

**UDF Usage**
```
=UDF_REGRESS_PREDICT(data, factorCols, resultCol, newVals)
```

#### InteractionEffects

成对交互效应检测。对每对因子构建含交互项的全模型，通过 F 检验比较全模型 vs 基准模型的残差平方和改善是否显著。支持数值/布尔/分类因子的任意组合。

**VBA Usage**
```vb
InteractionEffects(data, factorCols, resultCol, [hasHeader]) As Variant()
```

```vb
Dim interact As Variant
interact = InteractionEffects(dataRange, Array(1,2,3,4), 5)
' 输出含表头 "因子A/因子B/交互项数/F值/p值/显著"
' p < 0.05 -> 交互效应显著
```

> **如何解读输出**: 每行检验两个因子是否存在交互效应（综合效应是否不同于各自效应的简单相加）。**p < 0.05** 表示存在显著交互——这两个因子不是独立作用的。例如，温度×压力 p=0.003 意味着升温的效果取决于当前压力水平。p>0.05 的交互项表示两个因子可以独立优化。

**UDF Usage**
```
=UDF_REGRESS_INTERACT(data, factorCols, resultCol)
```

<a id="recipe-factor-optimization"></a>
### Recipe 7.3 — 因子优化

**场景**: 在低温和高压条件下搜索最大化产率的工艺参数组合，同时确保搜索空间可控。

**输入**: 同 Recipe 7.1 的数据 (温度 100-140, 压力 2.5-3.5, 时间 30-50)

`=UDF_REGRESS_OPTIMIZE(A1:D6, A1:C1, 4, "max", 3, 10)` ->

| 温度 | 压力 | 时间 | Predicted |
|------|------|------|-----------|
| 140 | 3.5 | 50 | 97.8 |
| 135.6 | 3.39 | 47.8 | 96.2 |
| 131.1 | 3.28 | 45.6 | 94.8 |

#### OptimizeFactors

网格搜索最优因子组合。对数值因子在 observed range 内生成 nSteps 个离散点，布尔因子枚举 {0,1}，分类因子枚举所有水平。搜索空间上限 MAX_GRID_COMBOS=200k。goal="max" 最大化，"min" 最小化，或传入数值做目标逼近。

**VBA Usage**
```vb
OptimizeFactors(data, factorCols, resultCol, [goal], [topN], [nSteps], [hasHeader]) As Variant()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| data | Range/Variant(,) | 训练数据 |
| factorCols | Variant | 因子列索引数组 |
| resultCol | Long | 结果列号 |
| goal | Variant | 可选. "max"/"min"/数值 (默认 "max") |
| topN | Long | 可选. 返回最优组合数 (默认 10) |
| nSteps | Long | 可选. 数值因子离散步数 (默认 10) |

```vb
Dim best As Variant
best = OptimizeFactors(dataRange, Array(1,2,3), 4, "max", 5, 10)
' 返回 top 5 最优组合
best = OptimizeFactors(dataRange, Array(1,2), 3, CDbl(7.5), 3, 20)
' 返回最接近 7.5 的 3 个组合
```

> **如何解读输出**: 返回 2D 数组，列包括：排名、预测值、每个因子的最佳设置。每行为一个候选组合，按优劣排序。
> - **goal="max"**: 按预测值降序排列（最佳在前）。
> - **goal="min"**: 按预测值升序排列。
> - **goal=数值**: 按与目标值的接近程度排列（最接近在前）。
> - **topN**: 返回多少组最优结果。建议查看多个候选——排名稍低的组合可能更实用。
> - **nSteps**: 越大搜索越精细，但运行时间指数增长。从 10 开始，必要时再增加。

**UDF Usage**
```
=UDF_REGRESS_OPTIMIZE(data, factorCols, resultCol, [goal], [topN], [nSteps])
```

#### FactorSweep

单因子 What-If 扫描。固定其他因子为基准值，扫描一个因子的变化对预测结果的影响。

**VBA Usage**
```vb
FactorSweep(model, factorIndex, fromVal, toVal, steps, [baseValues]) As Variant()
```

| 参数 | 类型 | 说明 |
|------|------|------|
| model | Object | LinearModelFit 返回的模型 |
| factorIndex | Long | 扫描因子序号 (在 factorCols 中的 1-based 序号) |
| fromVal | Double | 起始值 |
| toVal | Double | 结束值 |
| steps | Long | 扫描步数 |
| baseValues | Variant | 可选. 其他因子的基准值数组 |

```vb
Set model = LinearModelFit(dataRange, Array(1,2,3), 4)
Dim sweep As Variant
sweep = FactorSweep(model, 1, 100#, 140#, 11)
```

> **如何解读输出**: 返回 2D 数组，包含扫描因子的取值和对应的预测结果。绘制或查看"假设分析"曲线——当一个因子变化而其他因子保持不变时，预测结果如何变化。适用于可视化边际效应和寻找单个因子的最佳操作范围。

**UDF Usage**
```
=UDF_REGRESS_SWEEP(data, factorCols, resultCol, sweepFactorIdx, fromVal, toVal, steps)
```

---

> **依赖链**: RegressUtils -> LinearUtils + StatsUtils (TDist2T, FDistRT 用于 p 值计算; CorrelationMatrix 用于 UDF_REGRESS_CORREL)。导入时按 LinearUtils -> StatsUtils -> RegressUtils 顺序。

---

## Chapter 8: StringUtils — 字符串编码、校验、距离与格式化

字符串提取、编码转换、格式校验、编辑距离、随机生成与文本清洗。**模块**: `StringUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`ExtractBetween`](#extractbetween) | `(text, startDelim, endDelim)` | 提取分隔符之间文本 | String |
| [`ReverseString`](#reversestring) | `(text)` | 字符串反转 | String |
| [`CountSubstring`](#countsubstring) | `(text, substring, [caseSensitive])` | 统计子串出现次数 | Long |
| [`StartsWith`](#startswith) | `(text, prefix, [caseSensitive])` | 判断前缀是否匹配 | Boolean |
| [`EndsWith`](#endswith) | `(text, suffix, [caseSensitive])` | 判断后缀是否匹配 | Boolean |
| [`LeftOf`](#leftof) | `(text, delimiter, [nth])` | 取分隔符左侧文本 | String |
| [`RightOf`](#rightof) | `(text, delimiter, [nth], [fromRight])` | 取分隔符右侧文本 | String |
| [`TextJoin`](#textjoin) | `(delimiter, arr)` | 带分隔符连接数组 | String |
| [`NthWord`](#nthword) | `(text, n, [delimiter])` | 提取第 n 个词 | String |
| [`CommonPrefix`](#commonprefix) | `(a, b, [caseSensitive])` | 两字符串最长公共前缀 | String |
| [`PadLeft`](#padleft) | `(text, totalWidth, [padChar])` | 左侧填充至指定宽度 | String |
| [`PadRight`](#padleft) | `(text, totalWidth, [padChar])` | 右侧填充至指定宽度 | String |
| [`Repeat`](#repeat) | `(text, count)` | 重复字符串 n 次 | String |
| [`Truncate`](#truncate) | `(text, maxLen, [suffix])` | 截断并追加后缀 | String |
| [`NormalizeWhitespace`](#normalizewhitespace) | `(text)` | 折叠多余空白字符 | String |
| [`RemoveChars`](#removechars) | `(text, charsToRemove)` | 移除指定字符集 | String |
| [`KeepChars`](#removechars) | `(text, charsToKeep)` | 仅保留指定字符集 | String |
| [`ToTitleCase`](#totitlecase) | `(text)` | 智能标题大小写 (含 Mc/Mac/O') | String |
| [`RemoveDiacritics`](#removediacritics) | `(text)` | 移除变音符号 (Unicode NFD) | String |
| [`Slugify`](#slugify) | `(text)` | 生成 URL 友好 slug | String |
| [`IsNullOrEmpty`](#isnullorempty) | `(value)` | 判断是否 Null/Empty/ZLS | Boolean |
| [`IsNullOrWhitespace`](#isnullorempty) | `(value)` | 判断是否 Null/Empty/纯空白 | Boolean |
| [`IsEmail`](#isemail) | `(text)` | 邮箱格式校验 | Boolean |
| [`IsUrl`](#isemail) | `(text)` | URL 格式校验 | Boolean |
| [`Coalesce`](#coalesce) | `(ParamArray values())` | 返回首个非空值 (类 SQL COALESCE) | Variant |
| [`LevenshteinDistance`](#levenshteindistance) | `(s1, s2, [caseSensitive])` | 编辑距离 (Levenshtein) | Long |
| [`Soundex`](#soundex) | `(text)` | 发音哈希 (Soundex 算法) | String |
| [`URLEncode`](#urlencode) | `(text)` | URL 百分号编码 | String |
| [`URLDecode`](#urlencode) | `(text)` | URL 百分号解码 | String |
| [`Base64Encode`](#base64encode) | `(text)` | Base64 编码 | String |
| [`Base64Decode`](#base64encode) | `(text)` | Base64 解码 | String |
| [`HTMLEncode`](#htmlencode) | `(text)` | HTML 实体编码 | String |
| [`HTMLDecode`](#htmlencode) | `(text)` | HTML 实体解码 | String |
| [`UUID`](#uuid) | `()` | 生成 RFC 4122 v4 UUID | String |
| [`RandomString`](#randomstring) | `(length, [charset])` | 生成随机字符串 | String |

| Recipe | Functions Used |
|--------|---------------|
| [文本清洗](#recipe-text-cleaning) | `UDF_STR_NORMALIZEWHITESPACE`, `UDF_STR_REMOVECHARS`, `UDF_STR_TOTITLECASE` |
| [编码转换](#recipe-encoding) | `UDF_STR_BASE64ENCODE`, `UDF_STR_URLENCODE`, `UDF_STR_HTMLENCODE` |
| [模糊匹配](#recipe-fuzzy-match) | `UDF_STR_LEVENSHTEIN`, `UDF_STR_SOUNDEX` |

<a id="recipe-text-cleaning"></a>
### Recipe 8.1 — 文本清洗

**场景**: 从外部系统导入的数据含多余空格、特殊符号和大小写混乱，需标准化后入库。

`=UDF_STR_NORMALIZEWHITESPACE(UDF_STR_TOTITLECASE(UDF_STR_REMOVECHARS(A1, "™®")))`

<a id="removechars"></a>
<a id="keepchars"></a>
#### RemoveChars / KeepChars

RemoveChars 从字符串中移除指定字符集中的所有字符；KeepChars 仅保留指定字符集中的字符。

**VBA Usage**
```vb
RemoveChars(text, charsToRemove) As String
KeepChars(text, charsToKeep) As String
```

```vb
s = RemoveChars("AB-12/CD-34", "-/")   ' → "AB12CD34"
s = KeepChars("Tel: 555-1234", "0123456789")  ' → "5551234"
```

**UDF Usage**
```
=UDF_STR_REMOVECHARS(text, charsToRemove)
=UDF_STR_KEEPCHARS(text, charsToKeep)
```

#### NormalizeWhitespace

将连续空白字符 (空格/Tab/换行) 折叠为单个空格，并 Trim 首尾。

```vb
NormalizeWhitespace(text) As String
```

```vb
s = NormalizeWhitespace("  Hello   World  ")  ' → "Hello World"
```

**UDF Usage**
```
=UDF_STR_NORMALIZEWHITESPACE(text)
```

#### ToTitleCase

智能标题大小写转换：每个单词首字母大写，正确处理 Mc/Mac/O' 等前缀。

**VBA Usage**
```vb
ToTitleCase(text) As String
```

```vb
s = ToTitleCase("mcdonald o'brien macarthur")  ' → "McDonald O'Brien MacArthur"
```

**UDF Usage**
```
=UDF_STR_TOTITLECASE(text)
```

#### RemoveDiacritics

基于 Unicode NFD 分解移除变音符号 (如 é→e, ü→u, ñ→n)。

```vb
RemoveDiacritics(text) As String
```

```vb
s = RemoveDiacritics("crème brûlée")  ' → "creme brulee"
```

**UDF Usage**
```
=UDF_STR_REMOVEDIACRITICS(text)
```

#### Slugify

将文本转换为 URL 友好格式：小写 + 移除变音 + 非字母数字替换为连字符。

```vb
Slugify(text) As String
```

```vb
s = Slugify("Hello World!")  ' → "hello-world"
```

**UDF Usage**
```
=UDF_STR_SLUGIFY(text)
```

<a id="recipe-encoding"></a>
### Recipe 8.2 — 编码转换

**场景**: 需要将文本编码为 Base64 以便在 JSON/XML 中传输，或编码 URL 参数。

```
=UDF_STR_BASE64ENCODE(UDF_STR_URLENCODE(A1))
```

<a id="base64encode"></a>
<a id="base64decode"></a>
#### Base64Encode / Base64Decode

标准 Base64 编解码。编码输入为 ANSI/UTF-8 字节；解码返回原始字符串。

**VBA Usage**
```vb
Base64Encode(text) As String
Base64Decode(text) As String
```

```vb
s = Base64Encode("Hello")   ' → "SGVsbG8="
s = Base64Decode("SGVsbG8=") ' → "Hello"
```

**UDF Usage**
```
=UDF_STR_BASE64ENCODE(text)
=UDF_STR_BASE64DECODE(text)
```

<a id="urlencode"></a>
<a id="urldecode"></a>
#### URLEncode / URLDecode

标准 URL 百分号编解码。保留字母数字和 `-_.~`，其余字符编码为 `%XX`。

**VBA Usage**
```vb
URLEncode(text) As String
URLDecode(text) As String
```

```vb
s = URLEncode("Hello World")   ' → "Hello%20World"
s = URLDecode("Hello%20World") ' → "Hello World"
```

**UDF Usage**
```
=UDF_STR_URLENCODE(text)
=UDF_STR_URLDECODE(text)
```

<a id="htmlencode"></a>
<a id="htmldecode"></a>
#### HTMLEncode / HTMLDecode

HTML 实体编解码：`<>&"` 等特殊字符与 `&lt;` `&gt;` `&amp;` `&quot;` 互转。

**VBA Usage**
```vb
HTMLEncode(text) As String
HTMLDecode(text) As String
```

```vb
s = HTMLEncode("<div class='a'>")  ' → "&lt;div class=&apos;a&apos;&gt;"
s = HTMLDecode("&lt;p&gt;Hi&lt;/p&gt;") ' → "<p>Hi</p>"
```

**UDF Usage**
```
=UDF_STR_HTMLENCODE(text)
=UDF_STR_HTMLDECODE(text)
```

<a id="recipe-fuzzy-match"></a>
### Recipe 8.3 — 模糊匹配

**场景**: 需要在两个客户名单中根据姓名相似度进行匹配，容忍拼写差异。

```
=UDF_STR_LEVENSHTEIN(A1, B1)
=UDF_STR_SOUNDEX(A1)
```

#### LevenshteinDistance

计算两个字符串之间的编辑距离 (插入/删除/替换操作的最小次数)。0 表示完全相同。

```vb
LevenshteinDistance(s1, s2, [caseSensitive]) As Long
```

```vb
d = LevenshteinDistance("kitten", "sitting")  ' → 3
d = LevenshteinDistance("ABC", "abc", False)  ' → 0
```

**UDF Usage**
```
=UDF_STR_LEVENSHTEIN(text1, text2, [caseSensitive])
```

#### Soundex

根据英语发音生成 4 字符 Soundex 哈希码，用于模糊匹配发音相似的单词。

```vb
Soundex(text) As String
```

```vb
s = Soundex("Robert")  ' → "R163"
s = Soundex("Rupert")  ' → "R163" (与 Robert 相同发音组)
```

**UDF Usage**
```
=UDF_STR_SOUNDEX(text)
```

#### ExtractBetween

提取两个分隔符之间的文本 (不含分隔符本身)。

**VBA Usage**
```vb
ExtractBetween(text, startDelim, endDelim) As String
```

```vb
s = ExtractBetween("<title>Hello</title>", "<title>", "</title>") ' → "Hello"
```

**UDF Usage**
```
=UDF_STR_EXTRACTBETWEEN(text, startDelim, endDelim)
```

#### ReverseString

反转字符串顺序。

```vb
ReverseString(text) As String
```

```vb
s = ReverseString("ABC")  ' → "CBA"
```

**UDF Usage**
```
=UDF_STR_REVERSESTRING(text)
```

#### CountSubstring

统计子串在文本中出现的次数 (不重叠计数)。

```vb
CountSubstring(text, substring, [caseSensitive]) As Long
```

```vb
n = CountSubstring("banana", "an")  ' → 2
```

**UDF Usage**
```
=UDF_STR_COUNTSUBSTRING(text, substring, [caseSensitive])
```

<a id="startswith"></a>
<a id="endswith"></a>
#### StartsWith / EndsWith

判断字符串是否以指定前缀开始或后缀结束。

```vb
StartsWith(text, prefix, [caseSensitive]) As Boolean
EndsWith(text, suffix, [caseSensitive]) As Boolean
```

```vb
b = StartsWith("Hello World", "Hello")   ' → True
b = EndsWith("report.xlsx", ".xlsx")     ' → True
```

**UDF Usage**
```
=UDF_STR_STARTSWITH(text, prefix, [caseSensitive])
=UDF_STR_ENDSWITH(text, suffix, [caseSensitive])
```

<a id="leftof"></a>
<a id="rightof"></a>
#### LeftOf / RightOf

返回分隔符首次出现位置左侧或右侧的子串。未找到分隔符时返回空串。

```vb
LeftOf(text, delimiter, [nth]) As String
RightOf(text, delimiter, [nth], [fromRight]) As String
```

```vb
s = LeftOf("john.doe@example.com", "@")  ' → "john.doe"
s = RightOf("C:\Data\file.txt", "\")     ' → "Data\file.txt"
```

**UDF Usage**
```
=UDF_STR_LEFTOF(text, delimiter)
=UDF_STR_RIGHTOF(text, delimiter)
```

#### TextJoin

以分隔符连接数组，兼容旧版 Excel (无 TEXTJOIN)。

```vb
TextJoin(delimiter, arr) As String
```

```vb
s = TextJoin(", ", Array("A", "B", "C"))  ' → "A, B, C"
```

**UDF Usage**
```
=UDF_STR_TEXTJOIN(delimiter, arr)
```

#### NthWord

按分隔符提取第 n 个词 (1-based)。

```vb
NthWord(text, n, [delimiter]) As String
```

```vb
s = NthWord("apple,banana,cherry", 2, ",")  ' → "banana"
```

**UDF Usage**
```
=UDF_STR_NTHWORD(text, n, [delimiter])
```

#### CommonPrefix

返回两个字符串的最长公共前缀。默认不区分大小写。

```vb
CommonPrefix(a, b, [caseSensitive]) As String
```

```vb
s = CommonPrefix("flower", "flow")     ' → "flow"
s = CommonPrefix("flower", "flight")   ' → "fl"
s = CommonPrefix("abc", "ABC")         ' → "abc"  (不区分大小写)
```

**UDF Usage**
```
=UDF_STR_COMMONPREFIX(a, b, [caseSensitive])
```

<a id="padleft"></a>
<a id="padright"></a>
#### PadLeft / PadRight

用指定字符 (默认空格) 填充字符串至目标宽度。目标宽度小于文本长度时返回原文本。

```vb
PadLeft(text, totalWidth, [padChar]) As String
PadRight(text, totalWidth, [padChar]) As String
```

```vb
s = PadLeft("42", 5, "0")   ' → "00042"
s = PadRight("Name", 8)      ' → "Name    "
```

**UDF Usage**
```
=UDF_STR_PADLEFT(text, totalWidth, [padChar])
=UDF_STR_PADRIGHT(text, totalWidth, [padChar])
```

#### Truncate

截断文本至指定长度，超出部分以可选后缀 (默认 "...") 替代。

```vb
Truncate(text, maxLen, [suffix]) As String
```

```vb
s = Truncate("Hello World", 8)  ' → "Hello..."
```

**UDF Usage**
```
=UDF_STR_TRUNCATE(text, maxLen, [suffix])
```

<a id="isnullorempty"></a>
<a id="isnullorwhitespace"></a>
#### IsNullOrEmpty / IsNullOrWhitespace

空值判断：IsNullOrEmpty 检查 Null/Empty/零长度字符串；IsNullOrWhitespace 额外拒绝纯空白字符。

```vb
IsNullOrEmpty(value) As Boolean
IsNullOrWhitespace(value) As Boolean
```

```vb
b = IsNullOrEmpty("")        ' → True
b = IsNullOrWhitespace("  ") ' → True
b = IsNullOrEmpty("  ")      ' → False
```

**UDF Usage**
```
=UDF_STR_ISNULLOREMPTY(value)
=UDF_STR_ISNULLORWHITESPACE(value)
```

<a id="isemail"></a>
<a id="isurl"></a>
#### IsEmail / IsUrl

简单格式校验：IsEmail 检查 `@` 和域名结构；IsUrl 检查协议前缀和域名格式。

```vb
IsEmail(text) As Boolean
IsUrl(text) As Boolean
```

```vb
b = IsEmail("user@example.com")    ' → True
b = IsUrl("https://example.com")   ' → True
```

**UDF Usage**
```
=UDF_STR_ISEMAIL(text)
=UDF_STR_ISURL(text)
```

#### Coalesce

返回参数列表中第一个非空值 (不为 Null/Empty/ZLS)，类似 SQL COALESCE。

```vb
Coalesce(ParamArray values()) As Variant
```

```vb
v = Coalesce(Null, "", "Hello", "World")  ' → "Hello"
```

**UDF Usage**
```
=UDF_STR_COALESCE(value1, value2, ...)
```

<a id="uuid"></a>
<a id="randomstring"></a>
<a id="repeat"></a>
#### UUID / RandomString / Repeat

UUID 生成 RFC 4122 v4 UUID；RandomString 从指定字符集生成随机串；Repeat 重复字符串 n 次。

```vb
UUID() As String
RandomString(length, [charset]) As String
Repeat(text, count) As String
```

```vb
s = UUID()                            ' → "550e8400-e29b-41d4-a716-446655440000"
s = RandomString(8)                   ' → "aB3xK9mQ" (默认字母数字)
s = Repeat("ab", 3)                   ' → "ababab"
```

**UDF Usage**
```
=UDF_STR_UUID()
=UDF_STR_RANDOMSTRING(length, [charset])
=UDF_STR_REPEAT(text, count)
```

---

## Chapter 9: RegexUtils — 正则表达式匹配、替换、分割与捕获组

基于 VBScript.RegExp 的正则工具集，支持匹配测试、提取、替换、分割、计数与捕获组。**模块**: `RegexUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`RegexIsMatch`](#regexismatch) | `(text, pattern, [ignoreCase], [multiLine])` | 测试是否存在匹配 | Boolean |
| [`RegexExtract`](#regexextract) | `(text, pattern, [group], [ignoreCase])` | 提取首个匹配 (可指定捕获组) | String |
| `RegexFindInRange` | `(rng, pattern, [ignoreCase])` | 在 Range 中查找所有匹配 (VBA-only) | Range |
| [`RegexExtractAll`](#regexextractall) | `(text, pattern, [ignoreCase])` | 返回所有匹配数组 | String() |
| [`RegexExtractGroups`](#regexextractgroups) | `(text, pattern, [ignoreCase])` | 提取捕获组 → 2D 数组 | Variant(,) |
| [`RegexIsFullMatch`](#regexisfullmatch) | `(text, pattern, [ignoreCase])` | 测试是否完全匹配 | Boolean |
| [`RegexReplace`](#regexreplace) | `(text, pattern, replacement, [ignoreCase])` | 正则替换所有匹配 | String |
| [`RegexReplaceRange`](#regexreplacerange) | `(rng, pattern, replacement, [ignoreCase])` | 对区域执行替换 **Sub** | — |
| [`RegexSplit`](#regexsplit) | `(text, pattern, [ignoreCase])` | 正则分割文本 | String() |
| [`RegexSplitToRange`](#regexreplacerange) | `(text, pattern, destCell, [ignoreCase])` | 分割后输出到工作表 **Sub** | — |
| [`RegexCount`](#regexcount) | `(text, pattern, [ignoreCase])` | 统计匹配次数 | Long |
| [`RegexEscape`](#regexescape) | `(pattern)` | 转义正则特殊字符 | String |

| Recipe | Functions Used |
|--------|---------------|
| [正则提取](#recipe-regex-extract) | `UDF_REGEX_EXTRACT`, `UDF_REGEX_EXTRACTALL` |
| [正则替换](#recipe-regex-replace) | `UDF_REGEX_REPLACE` |
| [文本分割](#recipe-text-split) | `UDF_REGEX_SPLIT` |

<a id="recipe-regex-extract"></a>
### Recipe 9.1 — 正则提取

**场景**: 从自由文本中提取电话号码、邮箱地址或特定格式编号。

```
=UDF_REGEX_EXTRACT(A1, "\d{3}-\d{4}-\d{4}")
=UDF_REGEX_EXTRACTALL(A1, "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
```

#### RegexIsMatch

测试文本中是否存在正则模式匹配。返回 Boolean，不提取内容。

**VBA Usage**
```vb
RegexIsMatch(text, pattern, [ignoreCase], [multiLine]) As Boolean
```

```vb
b = RegexIsMatch("abc123", "\d+")        ' → True
b = RegexIsMatch("abc", "^\d+$")         ' → False
```

**UDF Usage**
```
=UDF_REGEX_ISMATCH(text, pattern, [ignoreCase], [multiLine])
```

#### RegexExtract

提取首个匹配内容。通过 `group` 参数可指定返回特定捕获组 (0=完整匹配)。

**VBA Usage**
```vb
RegexExtract(text, pattern, [group], [ignoreCase]) As String
```

```vb
s = RegexExtract("Phone: 555-1234", "\d{3}-\d{4}")  ' → "555-1234"
```

**UDF Usage**
```
=UDF_REGEX_EXTRACT(text, pattern, [group], [ignoreCase])
```

#### RegexExtractAll

返回所有非重叠匹配的字符串数组。

```vb
RegexExtractAll(text, pattern, [ignoreCase]) As String()
```

```vb
Dim matches() As String
matches = RegexExtractAll("a1 b2 c3", "\w\d")
' → {"a1","b2","c3"}
```

**UDF Usage**
```
=UDF_REGEX_EXTRACTALL(text, pattern, [ignoreCase])
```

#### RegexExtractGroups

提取所有匹配的捕获组，返回 2D 数组：每行一个匹配，每列一个捕获组。

```vb
RegexExtractGroups(text, pattern, [ignoreCase]) As Variant(,)
```

```vb
Dim groups As Variant
groups = RegexExtractGroups("Name: John, Age: 30", "(\w+): (\w+)")
' → {{"Name","John"},{"Age","30"}}
```

**UDF Usage**
```
=UDF_REGEX_EXTRACTGROUPS(text, pattern, [ignoreCase])
```

#### RegexIsFullMatch

测试整个字符串是否完全匹配正则模式 (等价于 `^pattern$`)。

```vb
RegexIsFullMatch(text, pattern, [ignoreCase]) As Boolean
```

```vb
b = RegexIsFullMatch("12345", "\d+")  ' → True
b = RegexIsFullMatch("abc123", "\d+") ' → False
```

**UDF Usage**
```
=UDF_REGEX_FULLMATCH(text, pattern, [ignoreCase])
```

<a id="recipe-regex-replace"></a>
### Recipe 9.2 — 正则替换

**场景**: 批量清洗数据，如移除 HTML 标签、统一日期格式、脱敏手机号。

```
=UDF_REGEX_REPLACE(A1, "<[^>]+>", "")
=UDF_REGEX_REPLACE(A1, "(\d{3})\d{4}(\d{4})", "$1****$2")
```

#### RegexReplace

替换文本中所有匹配正则模式的内容。支持 `$1`-`$9` 反向引用。

**VBA Usage**
```vb
RegexReplace(text, pattern, replacement, [ignoreCase]) As String
```

```vb
s = RegexReplace("2024-01-15", "(\d{4})-(\d{2})-(\d{2})", "$3/$2/$1")
' → "15/01/2024"
```

**UDF Usage**
```
=UDF_REGEX_REPLACE(text, pattern, replacement, [ignoreCase])
```

#### RegexReplaceRange (Sub)

对工作表区域中每个单元格执行正则替换，原地修改。

```vb
RegexReplaceRange(rng, pattern, replacement, [ignoreCase])
```

```vb
RegexReplaceRange Range("A1:A100"), "^\s+|\s+$", ""
```

<a id="recipe-text-split"></a>
### Recipe 9.3 — 文本分割

**场景**: 按多种分隔符拆分文本列，如逗号/分号/竖线混合分隔的数据。

```
=UDF_REGEX_SPLIT(A1, "[,;|]\s*")
```

#### RegexSplit

按正则模式分割文本，返回字符串数组。比 VBA `Split()` 更灵活 (支持多字符分隔符、空白容错)。

**VBA Usage**
```vb
RegexSplit(text, pattern, [ignoreCase]) As String()
```

```vb
Dim parts() As String
parts = RegexSplit("a,b; c|d", "[,;|]\s*")
' → {"a","b","c","d"}
```

**UDF Usage**
```
=UDF_REGEX_SPLIT(text, pattern, [ignoreCase])
```

#### RegexCount

统计正则模式在文本中的非重叠匹配次数。

```vb
RegexCount(text, pattern, [ignoreCase]) As Long
```

```vb
n = RegexCount("The fat cat sat on the mat", "\b\w{3}\b")  ' → 6
```

**UDF Usage**
```
=UDF_REGEX_COUNT(text, pattern, [ignoreCase])
```

#### RegexEscape

转义字符串中的正则特殊字符，使其作为字面量匹配。

```vb
RegexEscape(pattern) As String
```

```vb
s = RegexEscape("1+1=2?")  ' → "1\+1=2\?"
```

**UDF Usage**
```
=UDF_REGEX_ESCAPE(pattern)
```

---

## Chapter 10: JsonUtils — 纯 VBA JSON 解析、查询与序列化

纯 VBA 递归下降 JSON 解析器，无外部依赖。支持对象/数组/字符串/数字/布尔/null/嵌套和完整 Unicode (含代理对)。**模块**: `JsonUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`JsonParse`](#jsonparse) | `(jsonText)` | 解析 JSON 字符串 **仅 VBA** | Variant (Dictionary/Array/标量) |
| [`JsonGet`](#jsonget) | `(json, path)` | 按路径提取值 **仅 VBA** | Variant |
| [`JsonToRange`](#jsontorange) | `(json, destCell, [headers])` | JSON 对象数组输出到工作表 **Sub** | — |
| [`JsonGetKeys`](#jsongetkeys) | `(json)` | 列出对象所有键 | String() |
| [`JsonIsValid`](#jsonisvalid) | `(jsonText)` | 校验 JSON 合法性 (不抛错) | Boolean |
| [`JsonStringify`](#jsonstringify) | `(value)` | VBA Variant → JSON 字符串 **仅 VBA** | String |

| Recipe | Functions Used |
|--------|---------------|
| [JSON 路径提取](#recipe-json-path) | `UDF_JSON_GET` |
| [表格转 JSON](#recipe-table-to-json) | `UDF_JSON_STRINGIFY` |

<a id="recipe-json-path"></a>
### Recipe 10.1 — JSON 路径提取

**场景**: 从 API 返回的 JSON 响应中按路径提取嵌套字段值。

```
=UDF_JSON_GET(A1, "users[0].name")
=UDF_JSON_GET(A1, "data.items[2].price")
```

#### JsonParse

解析 JSON 字符串为 VBA 原生类型。对象 → Dictionary，数组 → Variant()，标量按类型转换。**仅 VBA** (Dictionary 无法直接返回到工作表)。对象键大小写敏感 (RFC 8259): `{"a":1,"A":2}` 保留两个键。

**VBA Usage**
```vb
JsonParse(jsonText) As Variant
```

```vb
Dim parsed As Variant
parsed = JsonParse("{""name"": ""John"", ""age"": 30}")
' parsed("name") → "John", parsed("age") → 30

' Unicode including surrogate pairs (emoji, rare CJK)
parsed = JsonParse("[""😀""]")
' parsed(0) → "😀"  (U+1F600, surrogate pair in VBA)
```

#### JsonGet

按点号/方括号路径从解析后的 JSON 中提取值。支持 `a.b`、`arr[0]`、`users[0].name`。**仅 VBA**。

**VBA Usage**
```vb
JsonGet(json, path) As Variant
```

```vb
Dim data As Variant: data = JsonParse("{""users"":[{""name"":""Alice""}]}")
v = JsonGet(data, "users[0].name")  ' → "Alice"
```

#### UDF_JSON_GET

工作表版 JsonGet。对象/数组返回占位文本 `[Object]` / `[Array]`，标量返回实际值。

**UDF Usage**
```
=UDF_JSON_GET(jsonText, path)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| jsonText | String | JSON 字符串 |
| path | String | 点号路径，如 `data.items[0].name` |

#### JsonGetKeys

返回 JSON 对象顶层所有键的字符串数组。

**VBA Usage**
```vb
JsonGetKeys(jsonText) As String()
```

```vb
Dim keys() As String
keys = JsonGetKeys("{""a"":1,""b"":2}")
' → {"a","b"}
```

**UDF Usage**
```
=UDF_JSON_KEYS(jsonText)
```

#### JsonIsValid

校验字符串是否为合法 JSON，不抛出异常。

```vb
JsonIsValid(jsonText) As Boolean
```

```vb
b = JsonIsValid("""hello""")  ' → True
b = JsonIsValid("{bad}")      ' → False
```

**UDF Usage**
```
=UDF_JSON_IS_VALID(text)
```

<a id="recipe-table-to-json"></a>
### Recipe 10.2 — 表格转 JSON

**场景**: 将工作表区域序列化为 JSON 字符串，用于 API 请求体或配置文件。

```
=UDF_JSON_STRINGIFY(range)
```

#### JsonStringify

将 VBA Variant (Dictionary/Array/标量) 序列化为 JSON 字符串。**仅 VBA**。

**VBA Usage**
```vb
JsonStringify(value) As String
```

```vb
s = JsonStringify(Array(1, 2, 3))  ' → "[1,2,3]"
```

#### JsonToRange (Sub)

将 JSON 对象数组输出到工作表，自动生成表头。

**VBA Usage**
```vb
JsonToRange(jsonText, destCell, [includeHeaders])
```

```vb
JsonToRange "[{""a"":1,""b"":2},{""a"":3,""b"":4}]", Range("A1")
```

---

## Chapter 11: XmlUtils — XML 解析与表格化

基于 MSXML2 (Windows 预装) 的 XML 工具。XPath 查询、属性提取、XML 转工作表。**模块**: `XmlUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|----------|-----------|-------------|---------|
| [`XmlValidate`](#xmlvalidate) | `(xml, [errDetail])` | 校验 XML 格式 | Boolean |
| [`XmlGet`](#xmlget) | `(xml, xpath)` | XPath 查询取值 | Variant |
| [`XmlGetAttr`](#xmlgetattr) | `(xml, xpath, attrName)` | 取属性值 | Variant |
| [`XmlToRange`](#xmltorange) | `(xml, rowXPath, [colNames])` | XML → 2D 数组 | Variant |

| Recipe | Functions Used |
|--------|---------------|
| [XPath 提取数据](#recipe-xpath-extract) | `UDF_XML_GET`, `XmlGetAttr` |
| [XML 转表格](#recipe-xml-to-table) | `UDF_XML_TABLE`, `XmlToRange` |

### Function Details

#### XmlValidate

校验 XML 格式。**VBA-only** 版本提供错误详情 (`errDetail` 输出参数)。

```vb
XmlValidate(xml, [errDetail]) As Boolean
```

```vb
Dim errInfo As Variant
If Not XmlValidate(badXml, errInfo) Then Debug.Print errInfo
```

**UDF Usage**
```
=UDF_XML_VALIDATE(xml)
```

`=UDF_XML_VALIDATE(A1)` → `TRUE` / `FALSE`

#### XmlGet

XPath 查询取节点文本。支持 `/a/b`、`//c`、`[@k='v']`、`[1]`。

```vb
XmlGet(xml, xpath) As Variant
```

```vb
XmlGet("<a><b>hello</b></a>", "/a/b")  ' → "hello"
```

**UDF Usage**
```
=UDF_XML_GET(xml, xpath)
```

`=UDF_XML_GET(A1, "/root/name")` → 节点文本

#### XmlGetAttr

取 XPath 节点的属性值。**VBA-only**。

```vb
XmlGetAttr(xml, xpath, attrName) As Variant
```

```vb
XmlGetAttr("<user id='007'/>", "/user", "id")  ' → "007"
```

#### XmlToRange

XML 重复节点 → 2D 数组（可直接写入 Range）。自动检测子元素列名。

```vb
XmlToRange(xml, rowXPath, [colNames]) As Variant
```

| Parameter | Type | Description |
|-----------|------|-------------|
| xml | String | XML 文本 |
| rowXPath | String | 行节点 XPath |
| colNames | Variant | (Optional) 列名数组；省略则自动检测 |

```vb
Dim arr As Variant
arr = XmlToRange(xml, "/rows/row", Array("colA", "colB"))
ws.Range("A1").Resize(UBound(arr,1), UBound(arr,2)).Value = arr
```

**UDF Usage**
```
=UDF_XML_TABLE(xml, rowXPath, [colNames])
```

`=UDF_XML_TABLE(A1, "/rows/row", {"a","b"})` → 2D 溢出数组

---


<a id="recipe-xpath-extract"></a>

### Recipe 11.1 — XPath 提取数据

**Scenario**: 从 XML 中按 XPath 提取指定节点或属性值。

**Input**
|   | A |
|---|---|
| 1 | `<user id="007"><name>Bond</name></user>` |

`=UDF_XML_GET(A1, "/user/name")` → `Bond`

**VBA Usage**
```vb
Dim v As Variant
v = XmlGet(xml, "/user/name")     ' → "Bond"
v = XmlGetAttr(xml, "/user", "id") ' → "007"
```

---

<a id="recipe-xml-to-table"></a>

### Recipe 11.2 — XML 转表格

**Scenario**: 将 XML 重复元素列表直接导入 Excel 工作表。

**Input**
|   | A |
|---|---|
| 1 | `<catalog><book><title>A</title><price>10</price></book><book><title>B</title><price>20</price></book></catalog>` |

`=UDF_XML_TABLE(A1, "/catalog/book", {"title","price"})` →

**Output**
|   | A | B |
|---|---|---|
| 1 | A | 10 |
| 2 | B | 20 |

**VBA Usage**
```vb
Dim arr As Variant
arr = XmlToRange(xml, "/catalog/book", Array("title", "price"))
ws.Range("A1").Resize(UBound(arr,1), UBound(arr,2)).Value = arr
```

---

## Chapter 12: DateTimeUtils — ISO 周、工作日、年龄与日期范围

日期信息提取、工作日计算、Unix 时间戳转换、年龄与节假日。**模块**: `DateTimeUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`ISOWeekNum`](#isoweeknum) | `(dt)` | ISO 8601 周数 (1-53) | Long |
| [`FirstDayOfMonth`](#firstdayofmonth) | `(dt)` | 当月第一天 | Date |
| [`LastDayOfMonth`](#lastdayofmonth) | `(dt)` | 当月最后一天 | Date |
| [`FirstDayOfQuarter`](#firstdayofquarter) | `(dt)` | 季度第一天 | Date |
| [`LastDayOfQuarter`](#lastdayofquarter) | `(dt)` | 季度最后一天 | Date |
| [`DaysInMonth`](#daysinmonth) | `(y, m)` | 当月天数 | Long |
| [`DaysInYear`](#daysinyear) | `(y)` | 当年天数 (365/366) | Long |
| [`DayOfYear`](#dayofyear) | `(dt)` | 一年中的第几天 (1-366) | Long |
| [`IsLeapYear`](#isleapyear) | `(year)` | 闰年判断 | Boolean |
| [`Quarter`](#quarter) | `(dt)` | 季度 (1-4) | Long |
| [`FiscalYear`](#fiscalyear) | `(dt, [startMonth])` | 财年 | Long |
| [`UnixToDate`](#unixtodate) | `(timestamp)` | Unix 时间戳 → Date | Date |
| [`DateToUnix`](#datetounix) | `(dt)` | Date → Unix 时间戳 | Double |
| [`IsWeekend`](#isweekend) | `(dt)` | 判断是否为周六/周日 | Boolean |
| [`IsHoliday`](#isholiday) | `(dt, [holidays])` | 判断是否为节假日 | Boolean |
| [`AddMonthsSafe`](#addmonthssafe) | `(dt, months)` | 安全加月 (月末自动截断) | Date |
| [`WorkdaysBetween`](#workdaysbetween) | `(startDate, endDate, [holidays])` | 两日期间工作日数 | Long |
| [`NextWorkday`](#nextworkday) | `(startDate, days, [holidays])` | n 个工作日后的日期 | Date |
| [`WeekOfMonth`](#weekofmonth) | `(dt)` | 当月第几周 (1-6) | Long |
| [`StartOfWeek`](#startofweek) | `(dt, [firstDayOfWeek])` | 所在周第一天 | Date |
| [`EndOfWeek`](#endofweek) | `(dt, [firstDayOfWeek])` | 所在周最后一天 | Date |
| [`DateRange`](#daterange) | `(startDate, endDate)` | 生成日期序列数组 | Date() |
| [`Age`](#age) | `(birthDate, [asOf])` | 详细年龄 (年/月/日) **仅 VBA** | Dictionary |
| [`AgeYears`](#ageyears) | `(birthDate, [refDate])` | 简单整数年龄 | Long |
| [`NthWeekday`](#nthweekday) | `(year, month, weekday, n)` | 某月第 n 个指定星期几 | Date |
| [`Easter`](#easter) | `(year)` | 复活节日期 (匿名算法) | Date |
| [`DateDiffParts`](#datediffparts) | `(d1, d2)` | 精确日期差 **仅 VBA** | Dictionary |

| Recipe | Functions Used |
|--------|---------------|
| [日期信息提取](#recipe-date-info) | `UDF_DT_ISOWEEKNUM`, `UDF_DT_QUARTER`, `UDF_DT_ISLEAPYEAR` |
| [工作日计算](#recipe-workday) | `UDF_DT_WORKDAYSBETWEEN`, `UDF_DT_NEXTWORKDAY` |
| [Unix 时间戳](#recipe-unix-timestamp) | `UDF_DT_UNIXTODATE`, `UDF_DT_DATETOUNIX` |

<a id="recipe-date-info"></a>
### Recipe 12.1 — 日期信息提取

**场景**: 从日期列提取 ISO 周数、季度和闰年标记，用于分组汇总。

```
=UDF_DT_ISOWEEKNUM(A1)
=UDF_DT_QUARTER(A1)
=UDF_DT_ISLEAPYEAR(YEAR(A1))
```

#### ISOWeekNum

返回 ISO 8601 周数 (1-53)。ISO 周定义：周一为周首日，第一周包含当年首个周四。

```vb
ISOWeekNum(dt) As Long
```

```vb
n = ISOWeekNum(#2025-01-01#)  ' → 1
```

**UDF Usage**
```
=UDF_DT_ISOWEEKNUM(date)
```

<a id="firstdayofmonth"></a>
<a id="lastdayofmonth"></a>
<a id="firstdayofquarter"></a>
<a id="lastdayofquarter"></a>
#### FirstDayOf/LastDayOf Month / Quarter

返回指定日期所在月或季度的第一天/最后一天。

```vb
FirstDayOfMonth(dt) As Date
LastDayOfMonth(dt) As Date
FirstDayOfQuarter(dt) As Date
LastDayOfQuarter(dt) As Date
```

```vb
d = LastDayOfMonth(#2024-02-15#)  ' → 2024-02-29
```

**UDF Usage**
```
=UDF_DT_FIRSTDAYOFMONTH(date)    =UDF_DT_LASTDAYOFMONTH(date)
=UDF_DT_FIRSTDAYOFQUARTER(date)  =UDF_DT_LASTDAYOFQUARTER(date)
```

<a id="daysinmonth"></a>
<a id="daysinyear"></a>
<a id="dayofyear"></a>
<a id="isleapyear"></a>
<a id="quarter"></a>
<a id="fiscalyear"></a>
#### DaysInMonth / DaysInYear / DayOfYear / IsLeapYear / Quarter / FiscalYear

日期组件提取函数。

```vb
DaysInMonth(y, m) As Long     ' 当月天数
DaysInYear(y) As Long         ' 当年天数
DayOfYear(dt) As Long         ' 第几天 (1-366)
IsLeapYear(year) As Boolean   ' 闰年判断
Quarter(dt) As Long           ' 季度 1-4
FiscalYear(dt, [startMonth]) As Long  ' 财年 (默认 1 月起始)
```

```vb
n = DaysInMonth(2024, 2)                 ' → 29
n = Quarter(#2024-07-15#)               ' → 3
n = FiscalYear(#2025-12-15#, 7)         ' → 2026 (7 月起)
```

**UDF Usage**
```
=UDF_DT_DAYSINMONTH(year, month)   =UDF_DT_DAYSINYEAR(year)
=UDF_DT_DAYOFYEAR(date)     =UDF_DT_ISLEAPYEAR(year)
=UDF_DT_QUARTER(date)       =UDF_DT_FISCALYEAR(date, [startMonth])  ' 可选 startMonth，默认 1
```

#### AgeYears

计算从出生日期到参考日期的整数年龄。

```vb
AgeYears(birthDate, [refDate]) As Long
```

```vb
n = AgeYears(#1990-06-15#, #2025-06-15#)  ' → 35
```

**UDF Usage**
```
=UDF_DT_AGEYEARS(birthDate, [refDate])
```

#### Age

计算从出生日期到参考日期的详细年龄 (年、月、日)。返回包含 "years"、"months"、"days" 键的 Dictionary。**仅 VBA**。

```vb
Age(birthDate, [asOf]) As Object
```

```vb
Dim a As Object
Set a = Age(#1990-06-15#, #2025-08-20#)
' a("years") → 35, a("months") → 2, a("days") → 5
```

<a id="weekofmonth"></a>
<a id="startofweek"></a>
<a id="endofweek"></a>
#### WeekOfMonth / StartOfWeek / EndOfWeek

WeekOfMonth 返回日期在当月的周序号 (1-6)；StartOfWeek/EndOfWeek 返回所在周的起止日期。

```vb
WeekOfMonth(dt) As Long
StartOfWeek(dt, [firstDayOfWeek]) As Date
EndOfWeek(dt, [firstDayOfWeek]) As Date
```

```vb
d = StartOfWeek(#2025-06-11#, vbMonday)  ' → 2025-06-09
```

**UDF Usage**
```
=UDF_DT_WEEKOFMONTH(date)
=UDF_DT_STARTOFWEEK(date, [firstDayOfWeek])
=UDF_DT_ENDOFWEEK(date, [firstDayOfWeek])
```

<a id="nthweekday"></a>
<a id="easter"></a>
#### NthWeekday / Easter

NthWeekday 返回指定年月中第 n 个指定星期几的日期；Easter 返回复活节日期。

```vb
NthWeekday(year, month, weekday, n) As Date
Easter(year) As Date
```

```vb
d = NthWeekday(2025, 6, vbMonday, 2)  ' → 2025-06-09 (6 月第 2 个周一)
d = Easter(2025)                        ' → 2025-04-20
```

**UDF Usage**
```
=UDF_DT_NTHWEEKDAY(year, month, weekday, n)
=UDF_DT_EASTER(year)
```

#### DateDiffParts

计算两个日期的精确差异，返回 Dictionary 含 years/months/days/totalDays。**仅 VBA**。

```vb
DateDiffParts(d1, d2) As Object  ' → Dictionary
```

```vb
Set diff = DateDiffParts(#2020-03-15#, #2025-06-11#)
' diff("years")=5, diff("months")=2, diff("days")=27
```

<a id="recipe-workday"></a>
### Recipe 12.2 — 工作日计算

**场景**: 计算项目工期 (工作日天数) 或确定 n 个工作日后的交付日期。

```
=UDF_DT_WORKDAYSBETWEEN(A1, B1, holidaysRange)
=UDF_DT_NEXTWORKDAY(A1, 10, holidaysRange)
```

<a id="workdaysbetween"></a>
<a id="nextworkday"></a>
#### WorkdaysBetween / NextWorkday

WorkdaysBetween 计算两日期间的工作日数 (周一至周五，含起止日)；NextWorkday 返回 n 个工作日后的日期。支持可选节假日列表。

```vb
WorkdaysBetween(startDate, endDate, [holidays]) As Long
NextWorkday(startDate, days, [holidays]) As Date
```

```vb
n = WorkdaysBetween(#2025-06-09#, #2025-06-13#)  ' → 5
d = NextWorkday(#2025-06-11#, 3)                  ' → 2025-06-16 (跳过周末)
```

**UDF Usage**
```
=UDF_DT_WORKDAYSBETWEEN(startDate, endDate, [holidays])
=UDF_DT_NEXTWORKDAY(startDate, days, [holidays])
```

<a id="isweekend"></a>
<a id="isholiday"></a>
#### IsWeekend / IsHoliday

IsWeekend 判断是否为周六或周日；IsHoliday 判断是否在指定节假日列表中。

```vb
IsWeekend(dt) As Boolean
IsHoliday(dt, [holidays]) As Boolean
```

**UDF Usage**
```
=UDF_DT_ISWEEKEND(date)
=UDF_DT_ISHOLIDAY(date, [holidays])
```

#### AddMonthsSafe

安全添加月份：若目标日期超出当月最大天数 (如 1 月 31 日 + 1 个月)，自动回退到月末。

```vb
AddMonthsSafe(dt, months) As Date
```

```vb
d = AddMonthsSafe(#2025-01-31#, 1)  ' → 2025-02-28
```

**UDF Usage**
```
=UDF_DT_ADDMONTHSSAFE(date, months)
```

#### DateRange

生成从起始日期到结束日期的所有日期数组。

```vb
DateRange(startDate, endDate) As Date()
```

**UDF Usage**
```
=UDF_DT_DATERANGE(startDate, endDate)
```

<a id="recipe-unix-timestamp"></a>
### Recipe 12.3 — Unix 时间戳

**场景**: 与外部系统 (API/数据库) 交换含 Unix 时间戳的数据。

```
=UDF_DT_UNIXTODATE(A1)
=UDF_DT_DATETOUNIX(A1)
```

<a id="unixtodate"></a>
<a id="datetounix"></a>
#### UnixToDate / DateToUnix

Unix 时间戳 (秒，UTC 1970-01-01 起) 与 VBA Date 类型互转。转换基于 UTC，不含时区偏移。

```vb
UnixToDate(timestamp) As Date
DateToUnix(dt) As Double
```

```vb
d = UnixToDate(1717977600#)   ' → 2025-06-10
ts = DateToUnix(#2025-06-10#) ' → 1717977600
```

**UDF Usage**
```
=UDF_DT_UNIXTODATE(timestamp)
=UDF_DT_DATETOUNIX(date)
```

---

## Chapter 13: RangeUtils — 区域导出、运算与命名管理

工作表区域定位、导出 (HTML/JSON/Markdown/CSV)、筛选、排序、去重与命名区域管理。**模块**: `RangeUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`LastRow`](#lastrow) | `([ws])` | 工作表最后数据行号 | Long |
| [`LastCol`](#lastrow) | `([ws])` | 工作表最后数据列号 | Long |
| [`FirstRow`](#lastrow) | `([ws])` | 数据首行号 | Long |
| [`FirstCol`](#lastrow) | `([ws])` | 数据首列号 | Long |
| [`ColLetter`](#colletter) | `(colNum)` | 列号 → 字母 | String |
| [`ColNumber`](#colletter) | `(colLetter)` | 字母 → 列号 | Long |
| [`UsedRangeEx`](#lastrow) | `([ws])` | 实际数据区域 | Range |
| [`GetCellAddress`](#getcelladdress) | `(row, col, [absolute])` | 构建单元格地址 | String |
| [`RangeToHTML`](#rangetohtml) | `(rng, [includeHeaders])` | 区域 → HTML 表格 | String |
| [`RangeToJSON`](#rangetohtml) | `(rng, [includeHeaders])` | 区域 → JSON 数组 | String |
| [`RangeToMarkdown`](#rangetohtml) | `(rng, [includeHeaders])` | 区域 → Markdown 表格 | String |
| [`ExportRangeToCSV`](#exportrangetocsv) | `(rng, filePath, [delimiter])` | 导出到 CSV 文件 **Sub** | — |
| [`FilterRangeToArray`](#filterrangetoarray) | `(rng, filterCol, operator, value)` | 条件筛选返回数组 | Variant(,) |
| [`FindAll`](#findall) | `(rng, what, [lookIn])` | 查找所有匹配单元格 | Range |
| [`MergeRanges`](#mergeranges) | `(rng1, rng2)` | 取区域并集 | Range |
| [`IntersectRanges`](#mergeranges) | `(rng1, rng2)` | 取区域交集 | Range |
| [`RangeExists`](#rangeexists) | `(rangeName)` | 判断引用是否为 #REF! | Boolean |
| [`NamedRangeExists`](#namedrange-operations) | `(name)` | 判断命名区域是否存在 | Boolean |
| [`NamedRangeAdd`](#namedrange-operations) | `(name, rng, [scope])` | 创建/更新命名区域 **Sub** | — |
| [`NamedRangeDelete`](#namedrange-operations) | `(name, [scope])` | 删除命名区域 **Sub** | — |
| [`IsRangeEmpty`](#israngeempty) | `(rng)` | 区域是否全空 | Boolean |
| [`CountVisible`](#israngeempty) | `(rng)` | 可见行数 | Long |
| [`RangeDiff`](#rangediff) | `(rng1, rng2)` | 比较区域差异 | Variant(,) |
| [`SafeText`](#safetext) | `(cell)` | 安全转换单元格值为字符串 | String |
| [`ValuesEqual`](#rangeexists) | `(v1, v2)` | 值相等比较 (容错) | Boolean |
| [`CopyRangeToSheet`](#lastrow) | `(arr, destCell)` | 数组 → 工作表 **Sub** | — |
| [`ClearRangeContents`](#clearrangecontents) | `(rng)` | 清除内容 **Sub** | — |
| [`ClearRangeFormats`](#clearrangecontents) | `(rng)` | 清除格式 **Sub** | — |
| [`RemoveDuplicatesRange`](#removeduplicatesrange) | `(rng, cols)` | 按列去重 **Sub** | — |
| [`SortRange`](#sortrange) | `(rng, keyCol, [ascending])` | 按列排序 **Sub** | — |

| Recipe | Functions Used |
|--------|---------------|
| [区域导出为 HTML](#recipe-range-export-html) | `UDF_RANGE_TOHTML` |
| [区域导出为 JSON](#recipe-range-export-json) | `UDF_RANGE_TOJSON` |
| [条件筛选](#recipe-conditional-filter) | `UDF_RANGE_FILTER` |

<a id="recipe-range-export-html"></a>
### Recipe 13.1 — 区域导出为 HTML

**场景**: 将数据表快速导出为 HTML 片段，用于邮件正文或网页嵌入。

```
=UDF_RANGE_TOHTML(A1:D10, TRUE)
```

<a id="rangetohtml"></a>
<a id="rangetojson"></a>
<a id="rangetomarkdown"></a>
#### RangeToHTML / RangeToJSON / RangeToMarkdown

将区域导出为 HTML 表格、JSON 数组或 Markdown 表格字符串。可选包含表头行。

**VBA Usage**
```vb
RangeToHTML(rng, [includeHeaders]) As String
RangeToJSON(rng, [includeHeaders]) As String
RangeToMarkdown(rng, [includeHeaders]) As String
```

```vb
s = RangeToHTML(Range("A1:C3"), True)
s = RangeToJSON(Range("A1:C3"), True)
s = RangeToMarkdown(Range("A1:C3"), True)
```

**UDF Usage**
```
=UDF_RANGE_TOHTML(range, [includeHeaders])
=UDF_RANGE_TOJSON(range, [includeHeaders])
=UDF_RANGE_TOMD(range, [includeHeaders])
```

#### ExportRangeToCSV (Sub)

将区域导出为 CSV 文件，支持自定义分隔符。

```vb
ExportRangeToCSV rng, filePath, [delimiter]
```

```vb
ExportRangeToCSV Range("A1:D100"), "C:\Data\export.csv", ";"
```

#### SafeText

安全获取单元格文本值，错误值 (`#N/A`, `#VALUE!` 等) 返回其显示文本。

```vb
SafeText(cell) As String
```

**UDF Usage**
```
=UDF_RANGE_SAFETEXT(cell)
```

#### GetCellAddress

构建指定行列的单元格地址字符串，支持绝对/相对引用。

```vb
GetCellAddress(row, col, [absolute]) As String
```

```vb
s = GetCellAddress(5, 3, True)  ' → "$C$5"
```

**UDF Usage**
```
=UDF_RANGE_CELLADDRESS(row, col, [absolute])
```

<a id="recipe-range-export-json"></a>
### Recipe 13.2 — 区域导出为 JSON

**场景**: 将表格数据序列化为 JSON，用于 Web API 数据交换。

```
=UDF_RANGE_TOJSON(A1:D10, TRUE)
```

<a id="lastrow"></a>
<a id="lastcol"></a>
<a id="firstrow"></a>
<a id="firstcol"></a>
#### LastRow / LastCol / FirstRow / FirstCol

返回工作表的数据边界行列号。LastRow/LastCol 从最后一行/列向上/左扫描。

```vb
LastRow([ws]) As Long
LastCol([ws]) As Long
FirstRow([ws]) As Long
FirstCol([ws]) As Long
```

**UDF Usage**
```
=UDF_RANGE_LASTROW([sheet])    =UDF_RANGE_LASTCOL([sheet])
=UDF_RANGE_FIRSTROW([sheet])   =UDF_RANGE_FIRSTCOL([sheet])
```

<a id="colletter"></a>
<a id="colnumber"></a>
#### ColLetter / ColNumber

列号与列字母互转：ColLetter(1) → "A"，ColNumber("Z") → 26。

```vb
ColLetter(colNum) As String
ColNumber(colLetter) As Long
```

**UDF Usage**
```
=UDF_RANGE_COL_LETTER(colNum)
=UDF_RANGE_COL_NUM(colLetter)
```

<a id="rangeexists"></a>
<a id="israngeempty"></a>
<a id="countvisible"></a>
#### RangeExists / IsRangeEmpty / CountVisible

RangeExists 判断引用有效性 (非 #REF!)；IsRangeEmpty 判断区域是否全部为空；CountVisible 统计可见行数。

```vb
RangeExists(rangeName) As Boolean
IsRangeEmpty(rng) As Boolean
CountVisible(rng) As Long
```

**UDF Usage**
```
=UDF_RANGE_EXISTS(range)
=UDF_RANGE_ISEMPTY(range)
=UDF_RANGE_COUNTVISIBLE(range)
```

#### NamedRange Operations

命名区域管理：NamedRangeAdd 创建或更新命名区域；NamedRangeDelete 删除命名区域；NamedRangeExists 检查是否存在。

```vb
NamedRangeAdd(name, rng, [scope])       ' Sub
NamedRangeDelete(name, [scope])         ' Sub
NamedRangeExists(name) As Boolean
```

```vb
NamedRangeAdd "MyData", Range("A1:D100")
NamedRangeExists("MyData")  ' → True
```

**UDF Usage**
```
=UDF_RANGE_NAMEDEXISTS(name)
```

<a id="recipe-conditional-filter"></a>
### Recipe 13.3 — 条件筛选

**场景**: 根据某列条件筛选区域数据，结果返回数组供进一步处理。

```
=UDF_RANGE_FILTER(A1:D100, 2, ">=", 80)
```

#### FilterRangeToArray

按列筛选区域，返回满足条件的行组成的 2D 数组。支持 `=`, `<`, `>`, `<=`, `>=`, `<>`, `contains`, `regex`。

**VBA Usage**
```vb
FilterRangeToArray(rng, filterCol, operator, value) As Variant(,)
```

```vb
Dim result As Variant
result = FilterRangeToArray(Range("A1:D100"), 2, ">=", 80)
```

**UDF Usage**
```
=UDF_RANGE_FILTER(range, filterCol, operator, value)
```

> **已知限制**: 旧版 Excel 中该 UDF 不产生完整溢出数组（需 Ctrl+Shift+Enter 数组公式或 Excel 365 动态数组）；建议改用 VBA 直接调用 `FilterRangeToArray`。

#### FindAll

在区域中查找所有匹配单元格，返回 Union Range。支持按值/公式/注释查找。

```vb
FindAll(rng, what, [lookIn]) As Range
```

```vb
Set found = FindAll(Range("A1:A100"), "Yes")
```

<a id="mergeranges"></a>
<a id="intersectranges"></a>
#### MergeRanges / IntersectRanges

MergeRanges 返回两区域的并集；IntersectRanges 返回交集。

```vb
MergeRanges(rng1, rng2) As Range
IntersectRanges(rng1, rng2) As Range
```

#### RangeDiff

比较两个区域的值差异，返回差异单元格的行列及内容对照表。

```vb
RangeDiff(rng1, rng2, [outTruncated]) As Range
```

<a id="clearrangecontents"></a>
<a id="clearrangeformats"></a>
<a id="removeduplicatesrange"></a>
<a id="sortrange"></a>
#### Sub Procedures (Clear/Sort/RemoveDuplicates/Copy)

原地操作类子程序：

```vb
ClearRangeContents rng           ' 清除内容
ClearRangeFormats rng            ' 清除格式
AutoFitRange rng                 ' 自动调整列宽/行高
SortRange rng, keyCol, [ascending]   ' 按列排序
RemoveDuplicatesRange rng, cols  ' 按指定列去重
CopyRangeToSheet arr, destCell   ' 数组写入工作表
```

```vb
RemoveDuplicatesRange Range("A1:D100"), Array(1, 2)  ' 按第 1、2 列去重
CopyRangeToSheet myArray, Range("A1")
```

---

## Chapter 14: FileSystemUtils — UTF-8 文件读写、文件夹遍历与驱动器信息

基于 Scripting.FileSystemObject 和 ADODB.Stream 的文件系统工具，支持 UTF-8 文本读写、文件/文件夹枚举、路径解析与驱动器信息。**模块**: `FileSystemUtils.bas`

> **⚠️ 路径安全限制**: 所有路径参数经 `ValidateSafePath` 检查：拒绝含 `..` 的目录穿越路径与 UNC 路径（`\\server\share` 及 `//server/share`）。已知限制：不解析符号链接/junction。注意失败语义：除 `ReadBinaryFile` 抛错外，多数函数在路径被拒时静默返回空值/`False`（如 `ReadTextFile` 返回 `""`、`DeleteFile`/`CopyFileSafe` 返回 `False`）——使用 UNC 或含 `..` 路径的调用方升级后会得到“假成功/空结果”，请先自查路径。

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`NormalizePath`](#normalizepath) | `(path)` | 统一分隔符为 `\` | String |
| [`PathCombine`](#normalizepath) | `(path1, path2)` | 安全拼接路径 | String |
| [`IsPathValid`](#ispathvalid) | `(path)` | 路径语法是否合法 | Boolean |
| [`FileExists`](#fileexists) | `(filePath)` | 判断文件是否存在 | Boolean |
| [`FolderExists`](#fileexists) | `(folderPath)` | 判断文件夹是否存在 | Boolean |
| [`GetFileName`](#getfilename) | `(path)` | 提取文件名 (含扩展名) | String |
| [`GetBaseName`](#getfilename) | `(path)` | 提取文件名 (不含扩展名) | String |
| [`GetExtension`](#getfilename) | `(path)` | 提取扩展名 | String |
| [`GetFolderPath`](#getfolderpath) | `(path)` | 提取父文件夹路径 | String |
| [`GetFileSize`](#getfilesize) | `(filePath)` | 文件大小 (字节) | Double |
| [`GetFileSizeFmt`](#getfilesize) | `(filePath)` | 格式化文件大小 | String |
| [`FileModified`](#filemodified) | `(filePath)` | 文件最后修改时间 | Date |
| [`ReadTextFile`](#readtextfile) | `(filePath, [encoding])` | 读取文本内容 | String |
| [`WriteTextFile`](#writetextfile) | `(filePath, text, [encoding])` | 写入文本文件 **Sub** | — |
| [`AppendTextFile`](#appendtextfile) | `(filePath, text, [encoding])` | 追加文本 **Sub** | — |
| [`ListFiles`](#listfiles) | `(folderPath, [pattern])` | 列出文件路径数组 | String() |
| [`ListFolders`](#listfiles) | `(folderPath)` | 列出子文件夹路径数组 | String() |
| [`EnsureFolder`](#ensurefolder) | `(folderPath)` | 递归创建目录 | Boolean |
| [`TempFileName`](#tempfilename) | `([prefix], [extension])` | 生成临时文件名 | String |
| [`GetTempFolder`](#tempfilename) | `()` | 系统临时文件夹路径 | String |
| [`GetSpecialFolder`](#getspecialfolder) | `(folderType)` | 特殊文件夹路径 | String |
| [`DeleteFile`](#deletefile) | `(filePath)` | 删除文件 | Boolean |
| [`CopyFileSafe`](#deletefile) | `(sourcePath, destPath)` | 安全复制 (自动创建目录) | Boolean |
| [`ReadBinaryFile`](#readbinaryfile) | `(filePath)` | 读取二进制文件 | Byte() |
| [`WriteBinaryFile`](#readbinaryfile) | `(filePath, data)` | 写入二进制文件 **Sub** | — |
| [`GetDriveInfo`](#getdriveinfo) | `(driveLetter)` | 驱动器信息 **仅 VBA** | Dictionary |
| [`CopyFolder`](#copyfolder) | `(source, dest)` | 递归复制文件夹 **Sub** | — |
| [`DeleteFolder`](#copyfolder) | `(source)` | 递归删除文件夹 **Sub** | — |

| Recipe | Functions Used |
|--------|---------------|
| [文本文件批量读取](#recipe-batch-file-read) | `UDF_FS_READTEXT`, `UDF_FS_LISTFILES` |
| [路径解析](#recipe-path-parse) | `UDF_FS_FILENAME`, `UDF_FS_BASENAME`, `UDF_FS_EXTENSION` |
| [文件信息](#recipe-file-info) | `UDF_FS_FILESIZE`, `UDF_FS_FILEMODIFIED`, `UDF_FS_FILEEXISTS` |

<a id="recipe-batch-file-read"></a>
### Recipe 14.1 — 文本文件批量读取

**场景**: 读取文件夹内所有日志文件，合并分析。

> **注意**: `UDF_FS_LISTFILES` 返回路径数组。批量读取需以动态数组公式输入 (Excel 365) 或 Ctrl+Shift+Enter (旧版)。读取单文件可直接用 `UDF_FS_READTEXT` 传入路径。

```
=UDF_FS_READTEXT(UDF_FS_LISTFILES("C:\Logs", "*.log"))
```

#### ReadTextFile

读取文本文件内容，默认 UTF-8 编码。支持的编码：`utf-8`, `ansi`, `unicode`。

**VBA Usage**
```vb
ReadTextFile(filePath, [encoding]) As String
```

```vb
content = ReadTextFile("C:\Data\config.json", "utf-8")
```

**UDF Usage**
```
=UDF_FS_READTEXT(filePath, [encoding])
```

<a id="writetextfile"></a>
<a id="appendtextfile"></a>
#### WriteTextFile / AppendTextFile (Sub)

写入或追加文本内容。默认 UTF-8 编码，包含 BOM。

```vb
WriteTextFile filePath, text, [encoding]
AppendTextFile filePath, text, [encoding]
```

```vb
WriteTextFile "C:\output.txt", "Hello World", "utf-8"
AppendTextFile "C:\output.txt", vbCrLf & "Line 2", "utf-8"
```

<a id="readbinaryfile"></a>
<a id="writebinaryfile"></a>
#### ReadBinaryFile / WriteBinaryFile

读取或写入文件的原始字节内容。

```vb
ReadBinaryFile(filePath) As Byte()
WriteBinaryFile filePath, data
```

```vb
bytes = ReadBinaryFile("C:\data.bin")
WriteBinaryFile "C:\output.bin", bytes
```

<a id="listfiles"></a>
<a id="listfolders"></a>
#### ListFiles / ListFolders

ListFiles 列出文件夹内匹配模式的文件路径数组；ListFolders 列出子文件夹路径数组。

**VBA Usage**
```vb
ListFiles(folderPath, [pattern]) As String()
ListFolders(folderPath) As String()
```

```vb
Dim files() As String
files = ListFiles("C:\Data", "*.csv")
```

**UDF Usage**
```
=UDF_FS_LISTFILES(folderPath, [pattern])
=UDF_FS_LISTFOLDERS(folderPath)
```

#### EnsureFolder

递归创建文件夹路径，已存在则跳过。返回是否成功。

```vb
EnsureFolder(folderPath) As Boolean
```

**UDF Usage**
```
=UDF_FS_ENSUREFOLDER(folderPath)
```

<a id="recipe-path-parse"></a>
### Recipe 14.2 — 路径解析

**场景**: 从完整文件路径中提取文件名、基本名或扩展名。

```
=UDF_FS_FILENAME(A1)
=UDF_FS_BASENAME(A1)
=UDF_FS_EXTENSION(A1)
=UDF_FS_FOLDERPATH(A1)
```

<a id="getfilename"></a>
<a id="getbasename"></a>
<a id="getextension"></a>
<a id="getfolderpath"></a>
#### GetFileName / GetBaseName / GetExtension / GetFolderPath

从完整路径中提取各部分。

```vb
GetFileName(path) As String     ' → "report.xlsx"
GetBaseName(path) As String     ' → "report"
GetExtension(path) As String    ' → ".xlsx" (含点号)
GetFolderPath(path) As String   ' → "C:\Data"
```

```vb
s = GetFileName("C:\Data\report.xlsx")    ' → "report.xlsx"
s = GetBaseName("C:\Data\report.xlsx")    ' → "report"
s = GetExtension("C:\Data\report.xlsx")   ' → ".xlsx"
```

**UDF Usage**
```
=UDF_FS_FILENAME(path)    =UDF_FS_BASENAME(path)
=UDF_FS_EXTENSION(path)   =UDF_FS_FOLDERPATH(path)
```

<a id="normalizepath"></a>
<a id="pathcombine"></a>
#### NormalizePath / PathCombine

NormalizePath 统一路径分隔符为反斜杠；PathCombine 安全拼接两段路径 (处理重复分隔符)。

```vb
NormalizePath(path) As String
PathCombine(path1, path2) As String
```

```vb
s = NormalizePath("C:/Data//files")       ' → "C:\Data\files"
s = PathCombine("C:\Data", "sub\file")    ' → "C:\Data\sub\file"
```

**UDF Usage**
```
=UDF_FS_NORMALIZEPATH(path)
=UDF_FS_PATHCOMBINE(path1, path2)
```

#### IsPathValid

检查路径字符串语法是否合法 (不含无效字符、格式正确)。

```vb
IsPathValid(path) As Boolean
```

**UDF Usage**
```
=UDF_FS_ISPATHVALID(path)
```

<a id="recipe-file-info"></a>
### Recipe 14.3 — 文件信息

**场景**: 列出文件清单并附带大小、修改时间和存在性信息。

```
=UDF_FS_FILESIZE(A1)
=UDF_FS_FILESIZEFMT(A1)
=UDF_FS_FILEMODIFIED(A1)
=UDF_FS_FILEEXISTS(A1)
```

<a id="getfilesize"></a>
<a id="getfilesizefmt"></a>
<a id="filemodified"></a>
<a id="fileexists"></a>
<a id="folderexists"></a>
#### GetFileSize / GetFileSizeFmt / FileModified / FileExists / FolderExists

文件信息查询函数。

```vb
GetFileSize(filePath) As Double          ' 字节数
GetFileSizeFmt(filePath) As String       ' 格式化如 "1.5 MB"
FileModified(filePath) As Date           ' 最后修改时间
FileExists(filePath) As Boolean          ' 文件是否存在
FolderExists(folderPath) As Boolean      ' 文件夹是否存在
```

```vb
n = GetFileSize("C:\Data\large.dat")          ' → 1572864
s = GetFileSizeFmt("C:\Data\large.dat")       ' → "1.50 MB"
d = FileModified("C:\Data\report.xlsx")       ' → 2025-06-10 14:30:00
```

**UDF Usage**
```
=UDF_FS_FILESIZE(path)       =UDF_FS_FILESIZEFMT(path)
=UDF_FS_FILEMODIFIED(path)   =UDF_FS_FILEEXISTS(path)
=UDF_FS_FOLDEREXISTS(path)
```

<a id="deletefile"></a>
<a id="copyfilesafe"></a>
#### DeleteFile / CopyFileSafe

DeleteFile 删除文件 (成功返回 True)；CopyFileSafe 复制文件并自动创建目标目录。

```vb
DeleteFile(filePath) As Boolean
CopyFileSafe(sourcePath, destPath) As Boolean
```

**UDF Usage**
```
=UDF_FS_DELETEFILE(path)
=UDF_FS_COPYFILE(source, dest)
```

<a id="tempfilename"></a>
<a id="gettempfolder"></a>
<a id="getspecialfolder"></a>
#### TempFileName / GetTempFolder / GetSpecialFolder

TempFileName 在临时目录生成唯一文件名；GetTempFolder 返回系统临时目录；GetSpecialFolder 返回桌面/文档等特殊文件夹。

```vb
TempFileName([prefix], [extension]) As String
GetTempFolder() As String
GetSpecialFolder(folderType) As String
```

```vb
s = GetSpecialFolder("Desktop")  ' → "C:\Users\...\Desktop"
s = GetSpecialFolder("MyDocuments") ' → "C:\Users\...\Documents"
```

**UDF Usage**
```
=UDF_FS_TEMPFILENAME([prefix], [extension])
=UDF_FS_TEMPFOLDER()
=UDF_FS_SPECIALFOLDER(folderType)
```

#### GetDriveInfo

返回指定驱动器信息 Dictionary，含 TotalSize/FreeSpace/FileSystem/VolumeName/IsReady。**仅 VBA**。

```vb
GetDriveInfo(driveLetter) As Object  ' → Dictionary
```

```vb
Set info = GetDriveInfo("C")
' info("TotalSize") → 256060514304
' info("FreeSpace") → 128030257152
```

<a id="copyfolder"></a>
<a id="deletefolder"></a>
#### CopyFolder / DeleteFolder (Sub)

递归复制或删除整个文件夹。

```vb
CopyFolder source, dest
DeleteFolder source
```

**UDF Usage**
```
=UDF_FS_COPYFOLDER(source, dest)
=UDF_FS_DELETEFOLDER(source)
```

---

## Chapter 15: PhyChemUtils — 分子量计算、单位换算与气体状态方程

理化计算工具：化学式分子量、单位换算 (体积/压力/温度/质量)、溶液稀释、理想气体与 Peng-Robinson 实际气体。**模块**: `PhyChemUtils.bas`

**Quick Reference**

| 函数 | 参数 | 说明 | 返回值 |
|------|------|------|--------|
| [`MolecularWeight`](#molecularweight) | `(formula)` | 化学式 → 摩尔质量 (g/mol) | Double |
| [`ConvertMass`](#convertmass) | `(val, fromUnit, [toUnit])` | 质量单位换算 | Double |
| [`ConvertVolume`](#convertvolume) | `(val, fromUnit, [toUnit])` | 体积单位换算 | Double |
| [`ConvertPressure`](#convertpressure) | `(val, fromUnit, [toUnit])` | 压力单位换算 | Double |
| [`ConvertTemperature`](#converttemperature) | `(val, fromUnit, [toUnit])` | 温度单位换算 | Double |
| [`ConvertStandard`](#convertstandard) | `(v, p, T, MW)` | 气体标态体积与质量 | Variant(,) |
| [`MassToMoles`](#masstomoles) | `(mass, molWeight)` | 质量 (g) → 物质的量 (mol) | Double |
| [`MolesToMass`](#masstomoles) | `(moles, molWeight)` | 物质的量 (mol) → 质量 (g) | Double |
| [`DilutionSolve`](#dilutionsolve) | `(c1, v1, c2, v2)` | C1V1=C2V2 求解器 (三缺一) | Double |
| [`IdealGasLaw`](#idealgaslaw) | `(P, V, n, T)` | PV=nRT 求解器 (三缺一) | Double |
| [`Density`](#density) | `(m, v, rho)` | 密度求解器 (三缺一) | Double |
| [`PercentYield`](#percentyield) | `(actual, theoretical)` | 产率百分比 | Double |
| [`CompressFactorPR`](#compressfactorpr) | `(P, T, Tc, Pc, omega)` | Peng-Robinson 压缩因子 Z | Double |
| [`CylinderStdVolume`](#cylinderstdvolume) | `(cylVol, fillP, fillT, gasName)` | 钢瓶标态体积 (含 Z 因子) | Double |
| [`CylinderStdVolumeFromMass`](#cylinderstdvolumefrommass) | `(netWt, gasFormula)` | 钢瓶标态体积 (已知净重) | Double |

| Recipe | Functions Used |
|--------|---------------|
| [分子量与稀释计算](#recipe-molweight-dilution) | `UDF_PC_MOLWEIGHT`, `UDF_PC_DILUTION` |
| [气体换算](#recipe-gas-conversion) | `UDF_PC_CONVERTSTANDARD`, `UDF_PC_STDVOLUME` |
| [单位换算](#recipe-unit-conversion) | `UDF_PC_CONVERTMASS`, `UDF_PC_CONVERTVOLUME`, `UDF_PC_CONVERTPRESSURE`, `UDF_PC_CONVERTTEMPERATURE` |

<a id="recipe-molweight-dilution"></a>
### Recipe 15.1 — 分子量与稀释计算

**场景**: 根据化学式自动计算摩尔质量，或按 C1V1=C2V2 求解稀释配比。

```
=UDF_PC_MOLWEIGHT("H2SO4")
=UDF_PC_DILUTION(10, 100, , 500)     ' 求解 v2 → 200 mL
```

#### MolecularWeight

根据化学式自动计算摩尔质量 (g/mol)。支持括号、嵌套和常见元素周期表 103 种元素。

**VBA Usage**
```vb
MolecularWeight(formula) As Variant
```

```vb
mw = MolecularWeight("H2O")        ' → 18.015
mw = MolecularWeight("Ca(OH)2")    ' → 74.093
mw = MolecularWeight("Na2SO4")       ' → 142.043
mw = MolecularWeight("Fe4[Fe(CN)6]3") ' → 859.24  (supports [] and {} brackets)
mw = MolecularWeight("CuSO4·5H2O")    ' → 249.69  (supports · and + hydrate separators)
```

**UDF Usage**
```
=UDF_PC_MOLWEIGHT(formula)
```

#### DilutionSolve

C1 * V1 = C2 * V2 求解器。四个参数中传入三个数值，待求解参数传 0 或省略。返回求解值。

**VBA Usage**
```vb
DilutionSolve(c1, v1, c2, v2) As Double
```

```vb
v2_needed = DilutionSolve(10, 100, 5, 0)  ' → 200
c2_final = DilutionSolve(10, 50, 0, 200)  ' → 2.5
```

**UDF Usage**
```
=UDF_PC_DILUTION(c1, v1, c2, v2)
```

<a id="masstomoles"></a>
<a id="molestomass"></a>
#### MassToMoles / MolesToMass

质量与物质的量互转。

```vb
MassToMoles(mass, molWeight) As Double
MolesToMass(moles, molWeight) As Double
```

单位约定：mass 单位 g，moles 单位 mol，molWeight 单位 g/mol。

```vb
n = MassToMoles(58.44, 58.44)  ' → 1 (mol NaCl)
m = MolesToMass(2, 18.015)     ' → 36.03 (g H2O)
```

**UDF Usage**
```
=UDF_PC_MASSTOMOLES(mass, molWeight)
=UDF_PC_MOLESTOMASS(moles, molWeight)
```

#### Density

密度求解器：通过 m/V/rho 中任意两个已知量计算第三个。

```vb
Density(m, v, rho) As Double
```

单位约定：mass 单位 g，volume 单位 mL，rho 单位 g/mL。

```vb
v = Density(100, 0, 2.5)  ' → 40 (体积 = 质量/密度)
```

**UDF Usage**
```
=UDF_PC_DENSITY(m, v, rho)
```

#### PercentYield

计算反应产率百分比：实际产量 / 理论产量 * 100。

```vb
PercentYield(actual, theoretical) As Double
```

```vb
y = PercentYield(8.5, 10)  ' → 85 (%)
```

**UDF Usage**
```
=UDF_PC_YIELD(actual, theoretical)
```

#### ConvertMass

#### ConvertVolume

#### ConvertPressure

#### ConvertTemperature

单位换算函数。每个函数接受 `(value, fromUnit, [toUnit])` 参数，省略目标单位时默认转换为 SI 单位。

支持的 toUnit 选项 (首参数 value 自动识别 fromUnit)：

| 函数 | 支持的 From 单位 |
|------|-----------------|
| ConvertMass | `"kg"`, `"g"`, `"mg"`, `"lb"`, `"oz"`, `"ton"` |
| ConvertVolume | `"m3"`, `"L"`, `"mL"`, `"gal"`, `"ft3"`, `"bbl"` |
| ConvertPressure | `"Pa"`, `"kPa"`, `"MPa"`, `"atm"`, `"bar"`, `"psi"`, `"mmHg"`, `"Torr"` |
| ConvertTemperature | `"C"`, `"K"`, `"F"` |

```vb
v = ConvertVolume(1, "gal", "L")          ' → 3.785
v = ConvertPressure(1, "atm", "kPa")      ' → 101.325
v = ConvertTemperature(100, "C", "F")     ' → 212
```

**UDF Usage**
```
=UDF_PC_CONVERTMASS(val, from, [to])
=UDF_PC_CONVERTVOLUME(val, from, [to])
=UDF_PC_CONVERTPRESSURE(val, from, [to])
=UDF_PC_CONVERTTEMPERATURE(val, from, [to])
```

<a id="recipe-gas-conversion"></a>
### Recipe 15.2 — 气体换算

**场景**: 将气体在特定条件下的体积换算为标准状态 (0°C, 1 atm) 体积。

```
=UDF_PC_CONVERTSTANDARD(volume, pressure, temperature, molWeight)
=UDF_PC_STDVOLUME(cylVol, fillP, fillT, gasName)
```

#### ConvertStandard

将气体在任意 (P, T) 条件下的体积换算为标态 (0°C, 1 atm) 下的体积和质量。返回含 V_std 和 mass 的 2D 数组。

**VBA Usage**
```vb
ConvertStandard(v, p, T, MW) As Variant(,)
```

| 参数 | 类型 | 说明 |
|------|------|------|
| v | Double | 体积 (L) |
| p | Double | 压力 (kPa) |
| T | Double | 温度 (°C) |
| MW | Double | 摩尔质量 (g/mol) |

```vb
Dim result As Variant
result = ConvertStandard(10, 202.65, 25, 28.01)
' result(0,0)="标况体积(L)", result(0,1)="标况质量(g)"
' result(1,0)=18.61, result(1,1)=23.27
```

**UDF Usage**
```
=UDF_PC_CONVERTSTANDARD(v, p, T, MW)
```

#### IdealGasLaw

PV = nRT 求解器。四个参数中传入三个数值，待求解参数传 Empty (或省略)。R = 8.314 J/(mol·K)。**P 单位 Pa，V 单位 m³，T 单位 K**。

```vb
IdealGasLaw(P, V, n, T) As Double
```

```vb
' 标准状态下 1 mol 理想气体: P=101325 Pa, V≈0.022414 m³, T=273.15 K
n = IdealGasLaw(101325, 0.022414, Empty, 273.15)  ' → ~1 (mol)

' 求解摩尔数: P=202650 Pa, V=0.01 m³, T=298 K → n≈0.818 mol
n = IdealGasLaw(202650, 0.01, Empty, 298)  ' → ~0.818
```

**UDF Usage**
```
=UDF_PC_IDEALGASLAW(P, V, n, T)
```

#### CompressFactorPR

Peng-Robinson 状态方程计算压缩因子 Z。需要临界温度 Tc (K)、临界压力 Pc (kPa) 和偏心因子 omega。

```vb
CompressFactorPR(P, T, Tc, Pc, omega) As Double
```

```vb
Z = CompressFactorPR(1013.25, 300, 304.2, 7380, 0.225)  ' CO2 在 1 atm, 27°C
```

**UDF Usage**
```
=UDF_PC_COMPRESS(P, T, Tc, Pc, omega)
=UDF_PC_ZFACTOR(P, T, gasName)
```

<a id="recipe-unit-conversion"></a>
### Recipe 15.3 — 单位换算

**场景**: 实验室和工程计算中频繁涉及质量、体积、压力、温度单位的换算，使用统一的换算函数确保一致性和可追溯性。

**输入**
|   | A | B | C |
|---|---|---|---|
| 1 | 数值 | 来源单位 | 目标单位 |
| 2 | 500 | g | kg |
| 3 | 1 | gal | L |
| 4 | 1 | atm | kPa |
| 5 | 25 | C | K |

`=UDF_PC_CONVERTMASS(A2, B2, C2)` → `0.5`

`=UDF_PC_CONVERTVOLUME(A3, B3, C3)` → `3.785`

`=UDF_PC_CONVERTPRESSURE(A4, B4, C4)` → `101.325`

`=UDF_PC_CONVERTTEMPERATURE(A5, B5, C5)` → `298.15`

<a id="cylinderstdvolume"></a>
<a id="cylinderstdvolumefrommass"></a>
#### CylinderStdVolume / CylinderStdVolumeFromMass

CylinderStdVolume 根据钢瓶容积、充装压力和温度计算标态体积 (自动包含 Z 因子修正)。CylinderStdVolumeFromMass 根据钢瓶净重和气体化学式计算标态体积。

```vb
CylinderStdVolume(cylVol, fillP, fillT, gasName) As Double
CylinderStdVolumeFromMass(netWt, gasFormula) As Double
```

| 参数 | 类型 | 说明 |
|------|------|------|
| cylVol | Double | 钢瓶容积 (L) |
| fillP | Double | 充装压力 (kPa) |
| fillT | Double | 充装温度 (°C) |
| gasName | String | 气体名称 (如 "CO2", "N2", "O2") |
| netWt | Double | 钢瓶净重 (kg) |
| gasFormula | String | 化学式 (如 "CO2", "N2") |

```vb
V_std = CylinderStdVolume(40, 15000, 25, "N2")          ' 40L N2 钢瓶标态体积
V_std = CylinderStdVolumeFromMass(25, "CO2")            ' 25kg CO2 钢瓶标态体积
```

**UDF Usage**
```
=UDF_PC_STDVOLUME(cylVol, fillP, fillT, gasName)
=UDF_PC_STDVOLMASS(netWt, gasFormula)
```

---

