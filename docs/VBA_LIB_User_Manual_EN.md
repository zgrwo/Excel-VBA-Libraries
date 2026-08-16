# Excel VBA Libraries — User Manual

<!-- last_updated: 2026-06-17 -->

> [User Manual (CN)](../rules/user-manual.md) | [API Reference](../rules/api-reference.md)

## Table of Contents

### Part 1: Data Processing

- **Chapter 1: ArrayUtils — Array Operations**
  - [Recipe 1.1 — Score Analysis](#recipe-score-analysis)
  - [Recipe 1.2 — Dynamic Filtering](#recipe-dynamic-filter)
  - [Recipe 1.3 — Data Sampling](#recipe-data-sampling)
- **Chapter 2: DictSetUtils — Dictionary & Set Operations**
  - [Recipe 2.1 — Set Operations](#recipe-set-operations)
  - [Recipe 2.2 — Frequency Statistics](#recipe-frequency-stats)
  - [Recipe 2.3 — Cartesian Product](#recipe-cartesian-product)
- **Chapter 3: PivotUtils — Data Reshaping**
  - [Recipe 3.1 — Cross Table to Detail](#recipe-cross-table-to-detail)
  - [Recipe 3.2 — Group Summarization](#recipe-group-summary)
  - [Recipe 3.3 — Cross Join](#recipe-cross-join)
- **Chapter 4: SqlUtils — SQL Queries**
  - [Recipe 4.1 — Worksheet Query](#recipe-worksheet-query)
  - [Recipe 4.2 — Multi-Table Join](#recipe-multi-table-join)
  - [Recipe 4.3 — Group Aggregation](#recipe-group-aggregation)

### Part 2: Mathematics & Statistics

- **Chapter 5: LinearUtils — Matrices & Linear Algebra**
  - [Recipe 5.1 — SVD Singular Value Decomposition](#recipe-svd)
  - [Recipe 5.2 — Linear System Solving](#recipe-linear-system)
  - [Recipe 5.3 — Polynomial Fitting](#recipe-polyfit)
- **Chapter 6: StatsUtils — Statistics & Distributions**
  - [Recipe 6.1 — Descriptive Statistics](#recipe-descriptive-stats)
  - [Recipe 6.2 — Hypothesis Testing](#recipe-hypothesis-test)
  - [Recipe 6.3 — Correlation Analysis](#recipe-correlation)
- **Chapter 7: RegressUtils — Regression & ANOVA**
  - [Recipe 7.1 — Factor Importance Analysis](#recipe-factor-importance)
  - [Recipe 7.2 — One-Way ANOVA](#recipe-one-way-anova)
  - [Recipe 7.3 — Factor Optimization](#recipe-factor-optimization)

### Part 3: Text & Data Formats

- **Chapter 8: StringUtils — String Processing**
  - [Recipe 8.1 — Text Cleaning](#recipe-text-cleaning)
  - [Recipe 8.2 — Encoding Conversion](#recipe-encoding)
  - [Recipe 8.3 — Fuzzy Matching](#recipe-fuzzy-match)
- **Chapter 9: RegexUtils — Regular Expressions**
  - [Recipe 9.1 — Regex Extraction](#recipe-regex-extract)
  - [Recipe 9.2 — Regex Replacement](#recipe-regex-replace)
  - [Recipe 9.3 — Text Splitting](#recipe-text-split)
- **Chapter 10: JsonUtils — JSON Processing**
  - [Recipe 10.1 — JSON Path Extraction](#recipe-json-path)
  - [Recipe 10.2 — Table to JSON](#recipe-table-to-json)
- **Chapter 11: XmlUtils — XML Parsing**
  - [Recipe 11.1 — XPath Data Extraction](#recipe-xpath-extract)
  - [Recipe 11.2 — XML to Worksheet](#recipe-xml-to-table)

### Part 4: Date & Excel/File

- **Chapter 12: DateTimeUtils — Date & Time**
  - [Recipe 12.1 — Date Info Extraction](#recipe-date-info)
  - [Recipe 12.2 — Workday Calculation](#recipe-workday)
  - [Recipe 12.3 — Unix Timestamps](#recipe-unix-timestamp)
- **Chapter 13: RangeUtils — Range Operations**
  - [Recipe 13.1 — Range Export to HTML](#recipe-range-export-html)
  - [Recipe 13.2 — Export Range to JSON](#recipe-range-export-json)
  - [Recipe 13.3 — Conditional Filter](#recipe-conditional-filter)
- **Chapter 14: FileSystemUtils — File I/O**
  - [Recipe 14.1 — Batch Text File Reading](#recipe-batch-file-read)
  - [Recipe 14.2 — Path Parsing](#recipe-path-parse)
  - [Recipe 14.3 — File Information](#recipe-file-info)

### Part 5: Physical Chemistry

- **Chapter 15: PhyChemUtils — Physical Chemistry**
  - [Recipe 15.1 — Molecular Weight & Dilution](#recipe-molweight-dilution)
  - [Recipe 15.2 — Gas Conversion](#recipe-gas-conversion)
  - [Recipe 15.3 — Unit Conversion](#recipe-unit-conversion)

---

## Chapter 1: ArrayUtils -- Array Operations

Array creation, sorting, filtering, slicing, aggregation, searching, and transformation toolkit. Output arrays default to 0-based; 1D/2D conversion follows VBA Range conventions (1-based). **Module**: `ArrayUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`ArrayDims`](#arraydims) | `(arr)` | Get array dimension count (0 for uninitialized/scalar) | Long |
| [`IsArray1D`](#isarray1d) | `(arr)` | Check if 1D array | Boolean |
| [`ArrayUnique`](#arrayunique) | `(arr)` | Dedup, preserving first-occurrence order | Variant() 0-based |
| [`ArraySort`](#arraysort) | `(arr, [ascending])` | QuickSort, O(n log n) | Variant() 0-based |
| [`ArrayFilterByValue`](#arrayfilterbyvalue) | `(arr, matchValue, [operator])` | Filter by value with comparison operator | Variant() 0-based |
| [`ArrayCountIf`](#arraycountif) | `(arr, matchValue, [operator])` | Count elements matching condition | Long |
| [`ArrayConcat`](#arrayconcat) | `(arr1, arr2)` | Concatenate two 1D arrays | Variant() 0-based |
| [`ArraySlice`](#arrayslice) | `(arr, [start], [cnt])` | Extract sub-array (Python-style indices) | Variant() 0-based |
| [`ArrayFlatten`](#arrayflatten) | `(arr)` | Flatten 2D array row-wise to 1D | Variant() 0-based |
| [`ArrayTranspose1D`](#arraytranspose1d) | `(arr, [asColumn])` | 1D to 2D column/row vector (1-based) | Variant(,) 1-based |
| [`ArrayFind`](#arrayfind) | `(arr, value, [caseSensitive])` | Find first index of element (0-based) | Variant (Long/-1) |
| [`ArrayContains`](#arraycontains) | `(arr, value, [caseSensitive])` | Check if array contains a value | Variant (Boolean) |
| [`ArrayLookup`](#arraylookup) | `(lookupArray, lookupValue, lookupCol, [returnCols], [matchType])` | In-memory multi-column lookup | Variant |
| [`ArrayShuffle`](#arrayshuffle) | `(arr)` | Fisher-Yates shuffle | Variant() 0-based |
| [`ArraySample`](#arraysample) | `(arr, n, [withReplacement])` | Random sampling (with/without replacement) | Variant() 0-based |
| [`LinSpace`](#linspace) | `(start, endVal, n)` | Generate n evenly spaced points | Double() 0-based |
| [`RangeFill`](#rangefill) | `(start, count, [stepSize])` | Generate arithmetic sequence | Double() 0-based |
| [`ArrayChunk`](#arraychunk) | `(arr, size)` | Split into fixed-size chunks, pad with Empty | Variant(,) 0-based |
| [`ArrayMin`](#arraymin) | `(arr)` | Numeric minimum (Empty if no valid values) | Variant (Double/Empty) |
| [`ArrayMax`](#arraymax) | `(arr)` | Numeric maximum (Empty if no valid values) | Variant (Double/Empty) |
| [`ArraySum`](#arraysum) | `(arr)` | Sum with Kahan compensated summation | Variant (Double) |
| [`ArrayToString`](#arraytostring) | `(arr, [delimiter])` | Join elements with delimiter | Variant (String) |
| [`ArrayReverse`](#arrayreverse) | `(arr)` | Reverse 1D array order | Variant() 0-based |
| [`ArrayGetRow`](#arraygetrow) | `(arr, row)` | Extract row from 2D array (0-based) | Variant() 0-based |
| [`ArrayGetCol`](#arraygetcol) | `(arr, col)` | Extract column from 2D array (0-based) | Variant() 0-based |
| [`ArrayTranspose2D`](#arraytranspose2d) | `(arr)` | Transpose 2D array (1-based) | Variant(,) 1-based |
| [`ArrayEqual`](#arrayequal) | `(arr1, arr2, [caseSensitive])` | Element-wise equality comparison | Variant (Boolean) |
| [`ArrayProduct`](#arrayproduct) | `(arr)` | Product of numeric elements | Variant (Double) |
| [`CumSum`](#cumsum) | `(arr)` | Cumulative sum | Double() 0-based |
| [`ArgSort`](#argsort) | `(arr, [ascending])` | Return sorted index array | Long() 0-based |
| [`ArrayAny`](#arrayany) | `(arr, matchValue, [operator])` | Any element satisfies condition? | Variant (Boolean) |
| [`ArrayAll`](#arrayall) | `(arr, matchValue, [operator])` | All elements satisfy condition? | Variant (Boolean) |

| Recipe | Functions Used |
|--------|---------------|
| [Score Analysis](#recipe-score-analysis) | `UDF_ARR_SORT`, `UDF_ARR_UNIQUE`, `UDF_ARR_COUNTIF`, `UDF_ARR_MIN`, `UDF_ARR_MAX`, `UDF_ARR_SUM` |
| [Dynamic Filtering](#recipe-dynamic-filter) | `UDF_ARR_FILTER`, `UDF_ARR_SLICE` |
| [Data Sampling](#recipe-data-sampling) | `UDF_ARR_SHUFFLE`, `UDF_ARR_SAMPLE` |

<a id="recipe-score-analysis"></a>

### Recipe 1.1 -- Score Analysis

**Scenario**: Given a student score sheet, deduplicate, sort, find min/max/average, and count passing students.

**Input**
|   | A | B |
|---|---|---|
| 1 | Name | Score |
| 2 | Alice | 85 |
| 3 | Bob | 92 |
| 4 | Charlie | 78 |
| 5 | Dave | 65 |
| 6 | Alice | 85 |
| 7 | Eve | 91 |

`=UDF_ARR_SORT(UDF_ARR_UNIQUE(B2:B7))` → `{65, 78, 85, 91, 92}`

`=UDF_ARR_MIN(B2:B7)` → `65`

`=UDF_ARR_MAX(B2:B7)` → `92`

`=UDF_ARR_SUM(B2:B7)` → `411`

`=UDF_ARR_COUNTIF(B2:B7, 60, ">=")` → `6`

#### ArraySort

1D array QuickSort, O(n log n), not guaranteed stable. Supports numeric, string, date. Null sorts first, Error sorts last.

**VBA Usage**
```vb
ArraySort(arr, [ascending]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array or Range |
| ascending | Boolean | Optional. True=ascending (default), False=descending |

```vb
result = ArraySort(Array(3, 1, 4, 2))
' result → {1, 2, 3, 4}
```

**UDF Usage**
```
=UDF_ARR_SORT(arr, [ascending])
```

| Parameter | Type | Description |
|------|------|------|
| arr | Range/Array | Data range |
| ascending | Boolean | Optional. Ascending=TRUE (default), Descending=FALSE |

**Input**
|   | A |
|---|---|
| 1 | 3 |
| 2 | 1 |
| 3 | 4 |
| 4 | 2 |

`=UDF_ARR_SORT(A1:A4, TRUE)` →

**Output**
|   | A |
|---|---|
| 1 | 1 |
| 2 | 2 |
| 3 | 3 |
| 4 | 4 |

#### ArrayUnique

1D array dedup, preserves first-occurrence order. Null merged as duplicates, Empty merged with 0 (since `IsNumeric(Empty)` = True).

**VBA Usage**
```vb
ArrayUnique(arr) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array or Range |

```vb
result = ArrayUnique(Array(1, 2, 2, 3, 1))
' result → {1, 2, 3}
```

**UDF Usage**
```
=UDF_ARR_UNIQUE(arr)
```

| Parameter | Type | Description |
|------|------|------|
| arr | Range | Data range |

#### ArrayFilterByValue

Filter 1D array by value. Supports comparison operators, regex, and contains.

**VBA Usage**
```vb
ArrayFilterByValue(arr, matchValue, [operator]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| matchValue | Variant | Value to match |
| operator | String | Optional. "=" (default), "<", ">", "<=", ">=", "<>", "contains", "regex" |

```vb
result = ArrayFilterByValue(Array(3, 1, 4, 1, 5), 3, ">")
' result → {4, 5}
```

**UDF Usage**
```
=UDF_ARR_FILTER(arr, matchValue, [operator])
```

| Parameter | Type | Description |
|------|------|------|
| arr | Range | Data range |
| matchValue | Value | Value to match |
| operator | String | Optional. Comparison operator; also supports "contains", "regex" |

#### ArrayCountIf

Count elements matching a condition. Same parameters as ArrayFilterByValue, but returns count instead of array.

**VBA Usage**
```vb
ArrayCountIf(arr, matchValue, [operator]) As Long
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| matchValue | Variant | Value to match |
| operator | String | Optional. Comparison operator, default "=" |

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

Numeric aggregation: minimum, maximum, sum. ArraySum uses Kahan compensated summation to reduce floating-point error. When no valid numeric values exist, ArrayMin/ArrayMax return Empty; ArraySum returns 0.

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

Concatenate two 1D arrays (or scalars). Scalars are auto-wrapped as single-element arrays.

**VBA Usage**
```vb
ArrayConcat(arr1, arr2) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr1 | Variant | First array or scalar |
| arr2 | Variant | Second array or scalar |

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

Extract a sub-array from a 1D array. `start` is 0-based index (negative counts from end), `cnt` is element count (-1 = to end).

**VBA Usage**
```vb
ArraySlice(arr, [start], [cnt]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| start | Long | Optional. Start index (0-based), default 0. Negative counts from end |
| cnt | Long | Optional. Number of elements, default -1 = to end |

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

Flatten a 2D array row-wise into a 1D array.

**VBA Usage**
```vb
ArrayFlatten(arr) As Variant
```

```vb
' data is 2×3 2D array: {{1,2,3},{4,5,6}}
result = ArrayFlatten(data)
' result → {1, 2, 3, 4, 5, 6}
```

**UDF Usage**
```
=UDF_ARR_FLATTEN(arr)
```

**Input**
|   | A | B | C |
|---|---|---|---|
| 1 | 1 | 2 | 3 |
| 2 | 4 | 5 | 6 |

`=UDF_ARR_FLATTEN(A1:C2)` → `{1,2,3,4,5,6}`

#### ArrayTranspose1D

Convert a 1D array to a 2D column vector or row vector (1-based output, compatible with VBA Range).

**VBA Usage**
```vb
ArrayTranspose1D(arr, [asColumn]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| asColumn | Boolean | Optional. True=column vector (default), False=row vector |

```vb
result = ArrayTranspose1D(Array(1, 2, 3), True)
' result → (1 To 3, 1 To 1): {{1},{2},{3}}
```

**UDF Usage**
```
=UDF_ARR_TRANSPOSE(arr, [asColumn])
```

#### ArrayFind

Find the first index of an element (0-based offset). Returns -1 if not found. Supports case-sensitive search.

**VBA Usage**
```vb
ArrayFind(arr, value, [caseSensitive]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| value | Variant | Value to find |
| caseSensitive | Boolean | Optional. Default False (case-insensitive) |

```vb
idx = ArrayFind(Array(10, 20, 30, 40), 30)
' idx → 2
```

**UDF Usage**
```
=UDF_ARR_FIND(arr, value, [caseSensitive])
```

#### ArrayContains

Check if an array contains a value. Delegates to ArrayFind; returns Boolean.

**VBA Usage**
```vb
ArrayContains(arr, value, [caseSensitive]) As Variant
```

```vb
found = ArrayContains(Array("Apple", "Banana", "Orange"), "Banana")
' found → True
```

**UDF Usage**
```
=UDF_ARR_CONTAINS(arr, value, [caseSensitive])
```

#### ArrayLookup

In-memory multi-column lookup (VLOOKUP replacement). Searches a 2D array for a value in a specific column and returns data from specified return columns.

**VBA Usage**
```vb
ArrayLookup(lookupArray, lookupValue, lookupCol, [returnCols], [matchType]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| lookupArray | Variant | 2D lookup array (with or without headers) |
| lookupValue | Variant | Value to look up |
| lookupCol | Long | Lookup column index (1-based) |
| returnCols | Variant | Optional. Return column number or array of numbers. Default: all columns |
| matchType | Long | Optional. 0=exact match (default), 1=approximate match (ascending) |

```vb
result = ArrayLookup(data, "Zhang", 1, Array(2, 3))
' result → 2D row array: columns 2 and 3 for Zhang
```

**UDF Usage**
```
=UDF_ARR_LOOKUP(lookupArray, lookupValue, lookupCol, [returnCols], [matchType])
```

#### ArrayShuffle

Fisher-Yates shuffle of a 1D array. Does not modify the original array.

**VBA Usage**
```vb
ArrayShuffle(arr) As Variant
```

```vb
result = ArrayShuffle(Array(1, 2, 3, 4, 5))
' result → random permutation, e.g., {3, 1, 5, 2, 4}
```

**UDF Usage**
```
=UDF_ARR_SHUFFLE(arr)
```

#### ArraySample

Random sampling from an array. Supports both without-replacement (default) and with-replacement modes.

**VBA Usage**
```vb
ArraySample(arr, n, [withReplacement]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| n | Long | Sample size |
| withReplacement | Boolean | Optional. False=without replacement (default), True=with replacement |

```vb
sample = ArraySample(Array(1, 2, 3, 4, 5, 6), 3)
' sample → 3 random elements, e.g., {4, 1, 5}
```

**UDF Usage**
```
=UDF_ARR_SAMPLE(arr, n, [withReplacement])
```

#### LinSpace

Generate n evenly spaced points, similar to NumPy `linspace`.

**VBA Usage**
```vb
LinSpace(start, endVal, n) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| start | Double | Start value |
| endVal | Double | End value |
| n | Long | Number of points (n ≥ 2) |

```vb
result = LinSpace(0, 1, 5)
' result → {0, 0.25, 0.5, 0.75, 1}
```

**UDF Usage**
```
=UDF_ARR_LINSPACE(start, endVal, n)
```

#### RangeFill

Generate an arithmetic sequence.

**VBA Usage**
```vb
RangeFill(start, count, [stepSize]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| start | Double | Start value |
| count | Long | Number of elements |
| stepSize | Double | Optional. Step size, default 1 |

```vb
result = RangeFill(0, 5, 2)
' result → {0, 2, 4, 6, 8}
```

**UDF Usage**
```
=UDF_ARR_RANGEFILL(start, count, [stepSize])
```

#### ArrayChunk

Split a 1D array into fixed-size chunks as a 2D array. Each row is a chunk; partial chunks are padded with Empty.

**VBA Usage**
```vb
ArrayChunk(arr, size) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| size | Long | Chunk size |

```vb
result = ArrayChunk(Array(1, 2, 3, 4, 5), 2)
' result → {{1, 2}, {3, 4}, {5, Empty}}  (3 rows × 2 cols)
```

**UDF Usage**
```
=UDF_ARR_CHUNK(arr, size)
```

#### ArrayToString

Join array elements into a string with a delimiter. Null/Error/Object values are safely handled.

**VBA Usage**
```vb
ArrayToString(arr, [delimiter]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| delimiter | String | Optional. Delimiter, default ", " |

```vb
s = ArrayToString(Array(1, 2, 3), "-")
' s → "1-2-3"
```

**UDF Usage**
```
=UDF_ARR_TOSTRING(arr, [delimiter])
```

#### ArrayReverse

Reverse the order of a 1D array.

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

Extract a specified row from a 2D array (0-based index).

**VBA Usage**
```vb
ArrayGetRow(arr, row) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 2D array |
| row | Long | Row index (0-based) |

```vb
result = ArrayGetRow(data, 1)  ' 2nd row
```

**UDF Usage**
```
=UDF_ARR_GETROW(arr, row)
```

#### ArrayGetCol

Extract a specified column from a 2D array (0-based index).

**VBA Usage**
```vb
ArrayGetCol(arr, col) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 2D array |
| col | Long | Column index (0-based) |

```vb
result = ArrayGetCol(data, 0)  ' 1st column
```

**UDF Usage**
```
=UDF_ARR_GETCOL(arr, col)
```

#### ArrayTranspose2D

Transpose a 2D array (1-based output). 1D input delegates to ArrayTranspose1D.

**VBA Usage**
```vb
ArrayTranspose2D(arr) As Variant
```

**UDF Usage**
```
=UDF_ARR_TRANSPOSE2D(arr)
```

**Input**
|   | A | B | C |
|---|---|---|---|
| 1 | 1 | 2 | 3 |
| 2 | 4 | 5 | 6 |

`=UDF_ARR_TRANSPOSE2D(A1:C2)` →

**Output**
|   | A | B |
|---|---|---|
| 1 | 1 | 4 |
| 2 | 2 | 5 |
| 3 | 3 | 6 |

#### ArrayEqual

Element-wise comparison of two 1D arrays. Supports scalar comparison.

**VBA Usage**
```vb
ArrayEqual(arr1, arr2, [caseSensitive]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr1 | Variant | First array or scalar |
| arr2 | Variant | Second array or scalar |
| caseSensitive | Boolean | Optional. Default False |

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

Product of all numeric elements. Returns 0 if no valid numeric values (compatible with Excel PRODUCT).

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

Cumulative sum. Returns an array of the same length where element i = sum of arr(0..i). Non-numeric elements are skipped.

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

Return the sorted index array (similar to NumPy `argsort`). result[0] is the original index of the smallest value.

**VBA Usage**
```vb
ArgSort(arr, [ascending]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| ascending | Boolean | Optional. Default True |

```vb
result = ArgSort(Array(30, 10, 20))
' result → {1, 2, 0}  (value 10 at index 1, 20 at index 2, 30 at index 0)
```

**UDF Usage**
```
=UDF_ARR_ARGSORT(arr, [ascending])
```

<a id="arrayany"></a>
<a id="arrayall"></a>
#### ArrayAny / ArrayAll

Check if any / all elements satisfy a condition.

**VBA Usage**
```vb
ArrayAny(arr, matchValue, [operator]) As Variant
ArrayAll(arr, matchValue, [operator]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array or scalar |
| matchValue | Variant | Value to match |
| operator | String | Optional. Comparison operator, default "=" |

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

Get the number of array dimensions. Returns 0 for uninitialized arrays or scalars.

**VBA Usage**
```vb
ArrayDims(arr) As Long
```

```vb
d = ArrayDims(Array(1, 2, 3))   ' d → 1
d = ArrayDims(Empty)             ' d → 0
```

#### IsArray1D

Check if a variable is a 1D array. Returns False for uninitialized arrays.

**VBA Usage**
```vb
IsArray1D(arr) As Boolean
```

```vb
b = IsArray1D(Array(1, 2))      ' b → True
b = IsArray1D("scalar")          ' b → False
```

### Recipe 1.2 -- Dynamic Filtering

<a id="recipe-dynamic-filter"></a>

**Scenario**: From a product list, filter products with prices below a threshold and take the top N.

**Input**
|   | A | B |
|---|---|---|
| 1 | Product | Price |
| 2 | Laptop | 5999 |
| 3 | Mouse | 199 |
| 4 | Keyboard | 499 |
| 5 | Monitor | 2299 |
| 6 | Headphones | 299 |
| 7 | Tablet | 3499 |

`=UDF_ARR_SLICE(UDF_ARR_FILTER(B2:B7, 3000, "<"), 0, 3)` → `{199, 499, 2299}`

### Recipe 1.3 -- Data Sampling

<a id="recipe-data-sampling"></a>

**Scenario**: Randomly select 3 students from a roster for an event, without replacement.

**Input**
|   | A |
|---|---|
| 1 | Name |
| 2 | Zhang |
| 3 | Li |
| 4 | Wang |
| 5 | Zhao |
| 6 | Qian |
| 7 | Sun |

`=UDF_ARR_SHUFFLE(A2:A7)` → randomly shuffled result

`=UDF_ARR_SAMPLE(A2:A7, 3)` → 3 random unique names

---

## Chapter 2: DictSetUtils -- Dictionary & Set Operations

Dictionary creation, merging, sorting, filtering, and set operations (union/intersection/difference/symmetric difference/subset/cartesian product). Functions returning Object are marked **VBA-only**. **Module**: `DictSetUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`DictMerge`](#dictmerge) | `(dict1, dict2, [overwrite])` | Merge two dictionaries | Object **VBA-only** |
| [`DictMergeSum`](#dictmergesum) | `(dict1, dict2)` | Merge dicts, sum values for same keys | Object **VBA-only** |
| [`DictInvert`](#dictinvert) | `(dict)` | Swap keys and values, skip invalid | Object **VBA-only** |
| [`DictKeys`](#dictkeys) | `(dict)` | Return array of all keys | Variant() |
| [`DictValues`](#dictvalues) | `(dict)` | Return array of all values | Variant() |
| [`DictTo2DArray`](#dictto2darray) | `(dict)` | Convert dict to 2D key-value array | Variant(,) 1-based |
| [`ArrayToDict`](#arraytodict) | `(arr, [keyFn])` | Convert 1D array to dictionary | Object **VBA-only** |
| [`DictFrom2DArray`](#dictfrom2darray) | `(arr, [colKey], [colValue])` | Convert two columns of 2D array to dict | Object **VBA-only** |
| [`CountFrequency`](#countfrequency) | `(arr)` | Frequency count, returns dict | Object **VBA-only** |
| [`GroupCount`](#groupcount) | `(arr)` | Frequency count, outputs 2D array directly | Variant(,) 1-based |
| [`DictPick`](#dictpick) | `(dict, keys)` | Extract sub-dict by key list | Object **VBA-only** |
| [`DictCount`](#dictcount) | `(dict)` | Dict entry count | Long |
| [`SetUnion`](#setunion) | `(arr1, arr2)` | Set union | Variant() 0-based |
| [`SetIntersect`](#setintersect) | `(arr1, arr2)` | Set intersection | Variant() 0-based |
| [`SetDifference`](#setdifference) | `(arr1, arr2)` | Set difference (A - B) | Variant() 0-based |
| [`SetSymDifference`](#setsymdifference) | `(arr1, arr2)` | Set symmetric difference (A Δ B) | Variant() 0-based |
| [`SetIsSubset`](#setissubset) | `(arr1, arr2)` | Check if arr1 is a subset of arr2 | Boolean |
| [`SetEqual`](#setequal) | `(arr1, arr2)` | Check if two sets are equal | Boolean |
| [`DictFilterByValue`](#dictfilterbyvalue) | `(dict, matchValue, [operator])` | Filter sub-dict by value condition | Object **VBA-only** |
| [`DictSortByKey`](#dictsortbykey) | `(dict, [ascending])` | Sort by key, returns 2D array | Variant(,) 1-based |
| [`DictSortByValue`](#dictsortbyvalue) | `(dict, [ascending])` | Sort by value, returns 2D array | Variant(,) 1-based |
| [`DictGetDefault`](#dictgetdefault) | `(dict, key, [defaultValue])` | Safe get, returns default if key missing | Variant |
| [`DictRenameKey`](#dictrenamekey) | `(dict, oldKey, newKey)` | Rename key, returns new dict | Object **VBA-only** |
| [`DictRemoveKeys`](#dictremovekeys) | `(dict, keys)` | Batch delete keys, returns new dict | Object **VBA-only** |
| [`DictClone`](#dictclone) | `(dict)` | Shallow clone dict | Object **VBA-only** |
| [`DictIsEmpty`](#dictisempty) | `(dict)` | Check if dict is empty | Boolean |
| [`DictTopN`](#dicttopn) | `(dict, n, [ascending])` | Top N by value, returns 2D array | Variant(,) 1-based |
| [`SetCartesianProduct`](#setcartesianproduct) | `(arrA, arrB)` | Cartesian product, returns 2D paired array | Variant(,) 1-based |

| Recipe | Functions Used |
|--------|---------------|
| [Set Operations](#recipe-set-operations) | `UDF_DICT_UNION`, `UDF_DICT_INTERSECT`, `UDF_DICT_DIFFERENCE` |
| [Frequency Statistics](#recipe-frequency-stats) | `UDF_DICT_GROUPCOUNT`, `CountFrequency` |
| [Cartesian Product](#recipe-cartesian-product) | `UDF_DICT_CARTESIAN` |

<a id="recipe-set-operations"></a>
### Recipe 2.1 -- Set Operations

**Scenario**: Two class rosters: find union (all students), intersection (common students), and difference (class A only).

**Input**
|   | A | B |
|---|---|---|
| 1 | Class 1 | Class 2 |
| 2 | Alice | Bob |
| 3 | Charlie | Alice |
| 4 | Dave | Eve |

`=UDF_DICT_UNION(A2:A4, B2:B4)` → `{Alice, Charlie, Dave, Bob, Eve}`

`=UDF_DICT_INTERSECT(A2:A4, B2:B4)` → `{Alice}`

`=UDF_DICT_DIFFERENCE(A2:A4, B2:B4)` → `{Charlie, Dave}`

#### SetUnion

Set union. Returns all unique elements from both 1D arrays.

**VBA Usage**
```vb
SetUnion(arr1, arr2) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr1 | Variant | First 1D array |
| arr2 | Variant | Second 1D array |

```vb
result = SetUnion(Array(1, 2, 3), Array(3, 4, 5))
' result → {1, 2, 3, 4, 5}
```

**UDF Usage**
```
=UDF_DICT_UNION(arr1, arr2)
```

| Parameter | Type | Description |
|------|------|------|
| arr1 | Range | First range |
| arr2 | Range | Second range |

#### SetIntersect

Set intersection. Returns elements present in both arrays.

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

Set difference (A - B). Returns elements in arr1 but not in arr2.

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

Set symmetric difference (A Δ B). Returns elements appearing in only one of the two arrays.

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

Check if arr1 is a subset of arr2. Empty/null sets are treated as subsets of any set.

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

Check if two sets are equal (order-independent).

**VBA Usage**
```vb
SetEqual(arr1, arr2) As Boolean
```

```vb
eq = SetEqual(Array(1, 2, 3), Array(3, 2, 1))
' eq → True (order-independent)
```

**UDF Usage**
```
=UDF_DICT_ISEQUAL(arr1, arr2)
```

### Recipe 2.2 -- Frequency Statistics

<a id="recipe-frequency-stats"></a>

**Scenario**: Count the frequency of each region in sales records.

**Input**
|   | A |
|---|---|
| 1 | region |
| 2 | Beijing |
| 3 | Shanghai |
| 4 | Beijing |
| 5 | Guangzhou |
| 6 | Shanghai |
| 7 | Beijing |

`=UDF_DICT_GROUPCOUNT(A2:A7)` →

**Output**
|   | A | B |
|---|---|---|
| 1 | Beijing | 3 |
| 2 | Shanghai | 2 |
| 3 | Guangzhou | 1 |

#### GroupCount

Frequency count, outputs 2D array directly (key, count). Skips Error and Null values.

**VBA Usage**
```vb
GroupCount(arr) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |

```vb
result = GroupCount(Array("A", "B", "A", "C", "B", "A"))
' result → {{"A", 3}, {"B", 2}, {"C", 1}}
```

**UDF Usage**
```
=UDF_DICT_GROUPCOUNT(arr)
```

#### CountFrequency

Frequency count, returns dictionary (key -> count). **VBA-only**, returns a Scripting.Dictionary.

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

Merge two dictionaries. When overwrite=False, duplicate keys preserve dict1 values; when overwrite=True, dict2 values overwrite.

**VBA Usage**
```vb
DictMerge(dict1, dict2, [overwrite]) As Object
```

| Parameter | Type | Description |
|------|------|------|
| dict1 | Object | First dictionary |
| dict2 | Object | Second dictionary |
| overwrite | Boolean | Optional. Defaults to False (preserve dict1) |

```vb
Dim merged As Object
Set merged = DictMerge(dict1, dict2, True)
```

#### DictMergeSum

Merge two dictionaries, sum numeric values for same keys. Non-numeric values preserve dict2 values.

**VBA Usage**
```vb
DictMergeSum(dict1, dict2) As Object
```

#### DictInvert

Swap keys and values. Skips Object/Error/Null/Array type values (cannot serve as dictionary keys). **VBA-only**.

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

Returns all keys/values of the dictionary as Variant arrays.

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

Convert dict to 2D array (1 To n, 1 To 2), column 1=keys, column 2=values.

**VBA Usage**
```vb
DictTo2DArray(dict) As Variant
```

#### ArrayToDict

Convert 1D array to dict. keyFn="value" (default) uses element values as keys; keyFn="index" uses indices as keys. **VBA-only**.

**VBA Usage**
```vb
ArrayToDict(arr, [keyFn]) As Object
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 1D array |
| keyFn | String | Optional. "value" (default) or "index" |

#### DictFrom2DArray

Convert two columns of 2D array to dict (colKey -> colValue). 1D array input auto-uses indices as keys. **VBA-only**.

**VBA Usage**
```vb
DictFrom2DArray(arr, [colKey], [colValue]) As Object
```

| Parameter | Type | Description |
|------|------|------|
| arr | Variant | 2D array |
| colKey | Long | Optional. key column (1-based), default 1 |
| colValue | Long | Optional. valuecolumn (1-based), default 2 |

#### DictPick

Extract sub-dictionary by specified key list. **VBA-only**.

**VBA Usage**
```vb
DictPick(dict, keys) As Object
```

#### DictCount

Dict entry count. Nothing returns 0.

**VBA Usage**
```vb
DictCount(dict) As Long
```

#### DictFilterByValue

Filter sub-dict by value condition. Supports "=", "<", ">", "<=", ">=", "<>", "contains", "regex". **VBA-only**.

**VBA Usage**
```vb
DictFilterByValue(dict, matchValue, [operator]) As Object
```

<a id="dictsortbykey"></a>
<a id="dictsortbyvalue"></a>
#### DictSortByKey / DictSortByValue

Sort by key/value, returns 2D arrays (1 To n, 1 To 2).

**VBA Usage**
```vb
DictSortByKey(dict, [ascending]) As Variant
DictSortByValue(dict, [ascending]) As Variant
```

```vb
sorted = DictSortByValue(dict, False)  ' Sort by value descending
' sorted(1, 1) = key of maximum, sorted(1, 2) = maximum value
```

#### DictGetDefault

Safe value retrieval. Returns default value when key does not exist; returns Empty when default value is not provided.

**VBA Usage**
```vb
DictGetDefault(dict, key, [defaultValue]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| dict | Object | dictionary |
| key | Variant | Key to look up |
| defaultValue | Variant | Optional. Default value when key does not exist |

```vb
val = DictGetDefault(dict, "missing", 0)
' val → 0 (key does not exist)
```

#### DictRenameKey

Rename a key in the dictionary, returns a new dictionary (original unchanged). New key is skipped if it already exists. **VBA-only**.

**VBA Usage**
```vb
DictRenameKey(dict, oldKey, newKey) As Object
```

#### DictRemoveKeys

Batch delete specified keys, returns a new dictionary (original unchanged). **VBA-only**.

**VBA Usage**
```vb
DictRemoveKeys(dict, keys) As Object
```

#### DictClone

Shallow clone dict. If values are object references, the clone and original dictionary point to the same objects. **VBA-only**.

**VBA Usage**
```vb
DictClone(dict) As Object
```

#### DictIsEmpty

Check if dictionary is Nothing or entry count is 0.

**VBA Usage**
```vb
DictIsEmpty(dict) As Boolean
```

#### DictTopN

Take top N entries by value (default: ascending=False returns top N by max value), returns 2D array.

**VBA Usage**
```vb
DictTopN(dict, n, [ascending]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| dict | Object | dictionary |
| n | Long | Number of entries to return |
| ascending | Boolean | Optional. Defaults to False (descending, maximum first) |

```vb
top = DictTopN(dict, 3, False)
' Returns top 3 entries by value descending
```

### Recipe 2.3 -- Cartesian Product

<a id="recipe-cartesian-product"></a>

**Scenario**: Generate all combinations of sizes and colors.

**Input**
|   | A | B |
|---|---|---|
| 1 | Dim | Color |
| 2 | S | Red |
| 3 | M | Blue |
| 4 | L | |

`=UDF_DICT_CARTESIAN(A2:A4, B2:B3)` →

**Output**
|   | A | B |
|---|---|---|
| 1 | S | Red |
| 2 | S | Blue |
| 3 | M | Red |
| 4 | M | Blue |
| 5 | L | Red |
| 6 | L | Blue |

#### SetCartesianProduct

Cartesian product of two 1D arrays, returns 2D array (1 To nA*nB, 1 To 2).

**VBA Usage**
```vb
SetCartesianProduct(arrA, arrB) As Variant
```

**UDF Usage**
```
=UDF_DICT_CARTESIAN(arr1, arr2)
```

---

## Chapter 3: PivotUtils -- Data Reshaping

Data pivot/unpivot, group aggregation, table filter/transpose, and cross join. Functions marked `(Sub)` write directly to the worksheet with no return value. **Module**: `PivotUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`RawConversion`](#rawconversion) | `(srcRange, valueColRef, colDimRef, rowDimRef, [keepBlank], [sortLabels])` | Cross-table pivot (long to wide) | Variant(,) |
| [`RawConversionToRange`](#rawconversiontorange) | `(... destCell, ...)` (Sub) | Pivot result written directly to worksheet | None |
| [`Unpivot`](#unpivot) | `(rng, valueCols, nameCol, valCol, destCell, [idColIndices])` (Sub) | Unpivot (wide to long) | None |
| [`GroupBy`](#groupby) | `(rng, groupCol, aggCol, [aggFunc], [destCell])` | Group aggregation (SUM/COUNT/AVG/MIN/MAX) | Variant(,) |
| [`SplitColumnToRows`](#splitcolumntorows) | `(rng, colIdx, delimiter, destCell)` (Sub) | Split column into multiple rows by delimiter | None |
| [`MergeColumns`](#mergecolumns) | `(rng, colIndices, delimiter, newColName, destCell, [dropOrigCols])` (Sub) | Merge multiple columns into one | None |
| [`FilterTable`](#filtertable) | `(rng, colIdx, op, value, destCell)` (Sub) | Filter table rows by condition (includes header) | None |
| [`TransposeTable`](#transposetable) | `(rng, destCell)` (Sub) | Table row/column transpose | None |
| [`VLookupArray`](#vlookuparray) | `(dataArray, lookupValue, lookupCol, [returnCol])` | In-memory VLOOKUP | Variant |
| [`CrossJoin`](#crossjoin) | `(rng1, rng2)` | Cross join two tables (Cartesian Product) | Variant(,) |

| Recipe | Functions Used |
|--------|---------------|
| [Cross Table to Detail](#recipe-cross-table-to-detail) | `UDF_PIVOT_CONVERT` |
| [Group Summarization](#recipe-group-summary) | `UDF_PIVOT_GROUPBY` |
| [Cross Join](#recipe-cross-join) | `UDF_PIVOT_CROSSJOIN` |

### Recipe 3.1 -- Cross Table to Detail

<a id="recipe-cross-table-to-detail"></a>

**Scenario**: Pivot raw sales data by product and month into a cross table, showing amount in value cells.

**Input**
|   | A | B | C |
|---|---|---|---|
| 1 | product | month | amount |
| 2 | Apple | 1Mo | 100 |
| 3 | Apple | 2Mo | 150 |
| 4 | Banana | 1Mo | 80 |
| 5 | Banana | 1Mo | 120 |
| 6 | Apple | 2Mo | 200 |

`=UDF_PIVOT_CONVERT(A1:C6, "amount", "month", "product")` →

**Output**
|   | A | B | C | D | E |
|---|---|---|---|---|---|---|---|
| 1 | product | 1MoA | 1MoB | 2MoA | 2MoB |
| 2 | Apple | 100 | | 150 | 200 |
| 3 | Banana | 80 | 120 | | |

> Column header suffixes A/B indicate multiple records at the same intersection, expanded in appearance order.

#### RawConversion

Data pivot (long table -> cross table). Rearranges data by row dimension, column dimension, and value column.

**VBA Usage**
```vb
RawConversion(srcRange, valueColRef, colDimRef, rowDimRef, [keepBlank], [sortLabels]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| srcRange | Range | Source data range (includes header row) |
| valueColRef | Variant | Value column header text or cell ref |
| colDimRef | Variant | Column dimension header text or cell ref |
| rowDimRef | Variant | Row dimension header text or cell ref |
| keepBlank | Boolean | Optional. Fill empty/null with "" (True, default) or Empty (False) |
| sortLabels | Boolean | Optional. Sort row/column labels alphabetically, defaults to False |

```vb
result = RawConversion(Range("A1:C100"), "amount", "month", "product", True, True)
```

**UDF Usage**
```
=UDF_PIVOT_CONVERT(srcRange, valueColRef, colDimRef, rowDimRef, [keepBlank], [sortLabels])
```

#### RawConversionToRange

Write pivot result directly to specified worksheet location. Parameters same as RawConversion, with additional destCell to specify output start cell.

**VBA Usage**
```vb
RawConversionToRange srcRange, valueColRef, colDimRef, rowDimRef, destCell, [keepBlank], [sortLabels]
```

```vb
RawConversionToRange Range("A1:C100"), "amount", "month", "product", Range("E1"), True, True
```

#### Unpivot

Unpivot (wide table -> long table). Converts multiple value columns into attribute name / attribute value columns, preserving ID columns.

**VBA Usage**
```vb
Unpivot rng, valueCols, nameCol, valCol, destCell, [idColIndices]
```

| Parameter | Type | Description |
|------|------|------|
| rng | Range | Source range (includes header row) |
| valueCols | Variant | Value column indices to unpivot (single or array) |
| nameCol | String | Header text for the attribute name column |
| valCol | String | Header text for the attribute value column |
| destCell | Range | Output start cell |
| idColIndices | Variant | Optional. ID column indices to keep unchanged (default: all columns not in valueCols) |

### Recipe 3.2 -- Group Summarization

<a id="recipe-group-summary"></a>

**Scenario**: Summarize sales amount by region, calculate sum and average.

**Input**
|   | A | B |
|---|---|---|
| 1 | region | sales amount |
| 2 | Beijing | 1000 |
| 3 | Shanghai | 800 |
| 4 | Beijing | 1200 |
| 5 | Guangzhou | 600 |
| 6 | Shanghai | 900 |

`=UDF_PIVOT_GROUPBY(A1:B6, 1, 2, "SUM")` →

**Output**
|   | A | B |
|---|---|---|
| 1 | region | SUM(sales amount) |
| 2 | Beijing | 2200 |
| 3 | Shanghai | 1700 |
| 4 | Guangzhou | 600 |

`=UDF_PIVOT_GROUPBY(A1:B6, 1, 2, "AVG")` →

**Output**
|   | A | B |
|---|---|---|
| 1 | region | AVG(sales amount) |
| 2 | Beijing | 1100 |
| 3 | Shanghai | 850 |
| 4 | Guangzhou | 600 |

#### GroupBy

Group aggregation. Supports SUM/COUNT/AVG/MIN/MAX. Skips Null/Error/non-numeric values.

**VBA Usage**
```vb
GroupBy(rng, groupCol, aggCol, [aggFunc], [destCell]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| rng | Range | Source range (includes header row) |
| groupCol | Long | Group column index (1-based) |
| aggCol | Long | Aggregation column index (1-based) |
| aggFunc | String | Optional. "SUM" (default), "COUNT", "AVG", "MIN", "MAX" |
| destCell | Range | Optional. If provided, simultaneously writes to worksheet |

```vb
result = GroupBy(Range("A1:B100"), 1, 2, "AVG")
```

**UDF Usage**
```
=UDF_PIVOT_GROUPBY(rng, groupCol, aggCol, [aggFunc])
```

#### VLookupArray

In-memory VLOOKUP. Searches by column in Range or 2D array, returns specified column values from matching rows. Case-insensitive.

**VBA Usage**
```vb
VLookupArray(dataArray, lookupValue, lookupCol, [returnCol]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| dataArray | Variant | Data range (Range) or 2D array |
| lookupValue | Variant | Value to look up |
| lookupCol | Long | Lookup column index (1-based) |
| returnCol | Long | Optional. Return column index (1-based), default 2 |

```vb
result = VLookupArray(Range("A1:B10"), "Alice", 1, 2)
```

**UDF Usage**
```
=UDF_PIVOT_VLOOKUP(dataArray, lookupValue, lookupCol, [returnCol])
```

#### SplitColumnToRows

Split column into multiple rows by delimiter. Each split element occupies its own row; other columns copy original values.

**VBA Usage**
```vb
SplitColumnToRows rng, colIdx, delimiter, destCell
```

| Parameter | Type | Description |
|------|------|------|
| rng | Range | Source range (includes header row) |
| colIdx | Long | Column index to split (1-based) |
| delimiter | String | delimiter |
| destCell | Range | Output start cell |

#### MergeColumns

Merge multiple columns into one. Supports preserving or dropping original columns.

**VBA Usage**
```vb
MergeColumns rng, colIndices, delimiter, newColName, destCell, [dropOrigCols]
```

| Parameter | Type | Description |
|------|------|------|
| rng | Range | Source range |
| colIndices | Variant | Column indices to merge (single or array) |
| delimiter | String | delimiter |
| newColName | String | New column header |
| destCell | Range | Output start cell |
| dropOrigCols | Boolean | Optional. Whether to drop original columns, defaults to False |

#### FilterTable

Filter table by condition (preserves header row). Supports "=", "<", ">", "<=", ">=", "<>", "contains".

**VBA Usage**
```vb
FilterTable rng, colIdx, op, value, destCell
```

| Parameter | Type | Description |
|------|------|------|
| rng | Range | Source range (includes header row) |
| colIdx | Long | Filter column index (1-based) |
| op | String | Comparison operator |
| value | Variant | Comparison value |
| destCell | Range | Output start cell |

```vb
FilterTable Range("A1:C100"), 3, ">", 60, Range("E1")
```

#### TransposeTable

Table row/column transpose (swap rows and columns).

**VBA Usage**
```vb
TransposeTable rng, destCell
```

### Recipe 3.3 -- Cross Join

<a id="recipe-cross-join"></a>

**Scenario**: Generate all combinations of students and subjects (course scheduling table).

**Input**
|   | A | B |
|---|---|---|
| 1 | student | Subject |
| 2 | Alice | Math |
| 3 | Bob | English |

`=UDF_PIVOT_CROSSJOIN(A1:A3, B1:B3)` →

**Output**
|   | A | B |
|---|---|---|
| 1 | student | Subject |
| 2 | Alice | Math |
| 3 | Alice | English |
| 4 | Bob | Math |
| 5 | Bob | English |

#### CrossJoin

Cross Join (Cartesian Product) of two tables. Output row count = 1 + (rng1 data rows) * (rng2 data rows).

**VBA Usage**
```vb
CrossJoin(rng1, rng2) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| rng1 | Range | First range (includes header row) |
| rng2 | Range | Second range (includes header row) |

**UDF Usage**
```
=UDF_PIVOT_CROSSJOIN(rng1, rng2)
```

> **Known Limitation**: When used as a worksheet formula, the spill range may be truncated to the input row count; call `CrossJoin` directly from VBA for the full Cartesian product.

---

## Chapter 4: SqlUtils -- SQL Queries

SQL queries on Excel worksheets via ADODB (ACE/Jet OLEDB). Supports SELECT/JOIN/GROUP BY/WHERE/ORDER BY. **Module**: `SqlUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`CloseSqlCache`](#closesqlcache) | `()` (Sub) | Release cached database connection | None |
| [`SqlGetConnection`](#sqlgetconnection) | `([filePath], [outOk])` | Get ADODB connection (cached reuse) | Object **VBA-only** |
| [`SqlExecute`](#sqlexecute) | `(sql, [filePath], [includeHeader], outOk, [outErrorMsg])` | Execute any SQL, returns 2D array | Variant() **VBA-only** |
| [`SqlQuery`](#sqlquery) | `(selectClause, fromClause, [whereClause], [orderByClause], [filePath], [outOk])` | SELECT query (clause-based params) | Variant() |
| [`SqlJoin`](#sqljoin) | `(table1, table2, joinOn, [joinType], [selectCols], [filePath], [outOk])` | Two-table JOIN | Variant() |
| [`SqlGroupBy`](#sqlgroupby) | `(tableName, groupCols, aggExprs, [whereClause], [filePath], [outOk])` | Group aggregation | Variant() |
| [`SqlListSheets`](#sqllistsheets) | `([filePath], [outOk])` | List all worksheet names in workbook | Variant() |
| [`SqlListColumns`](#sqllistcolumns) | `(tableName, [filePath], [outOk])` | List column names and ordinals for a table | Variant() |
| [`SqlListTables`](#sqllisttables) | `([filePath], [outOk])` | List all available data sources (tables/named ranges) | Variant() |
| [`SqlRangeQuery`](#sqlrangequery) | `(sql, rng, [tableAlias], [outOk])` | Query a Range directly (no saved workbook needed) | Variant() |

| Recipe | Functions Used |
|--------|---------------|
| [Worksheet Query](#recipe-worksheet-query) | `UDF_SQL_QUERY` |
| [Multi-Table Join](#recipe-multi-table-join) | `UDF_SQL_JOIN` |
| [Group aggregation](#recipe-group-aggregation) | `UDF_SQL_GROUPBY` |

### Recipe 4.1 -- Worksheet Query

<a id="recipe-worksheet-query"></a>

**Scenario**: From the student score table, query students with scores above 80, sorted by score descending.

**Input** (worksheet name: `Sheet1`)

|   | A | B | C |
|---|---|---|---|
| 1 | student | Subject | score |
| 2 | Alice | Math | 85 |
| 3 | Bob | English | 92 |
| 4 | Charlie | Math | 78 |
| 5 | Dave | Chinese | 88 |

`=UDF_SQL_QUERY("SELECT * FROM [Sheet1$] WHERE score > 80 ORDER BY score DESC")` →

**Output**
|   | A | B | C |
|---|---|---|---|
| 1 | student | Subject | score |
| 2 | Bob | English | 92 |
| 3 | Dave | Chinese | 88 |
| 4 | Alice | Math | 85 |

#### SqlExecute

Execute any SQL statement, returns 2D Variant array (includes column name header row). **VBA-only** (indicates success/failure via outOk parameter).

**VBA Usage**
```vb
SqlExecute(sql, [filePath], [includeHeader], outOk, [outErrorMsg]) As Variant()
```

| Parameter | Type | Description |
|------|------|------|
| sql | String | SQL statement |
| filePath | String | Optional. Workbook path, defaults to current workbook |
| includeHeader | Boolean | Optional. Whether to include column name header row, defaults to True |
| outOk | Boolean | Output. Returns True on success |

```vb
Dim ok As Boolean
result = SqlExecute("SELECT * FROM [Sheet1$] WHERE Age > 30", , , ok)
If ok Then Range("A1").Resize(UBound(result,1), UBound(result,2)).Value = result
```

**UDF Usage**
```
=UDF_SQL_QUERY(sql, [filePath])
```

| Parameter | Type | Description |
|------|------|------|
| sql | String | Complete SQL statement |
| filePath | String | Optional. Workbook path, defaults to current workbook |

#### SqlQuery

SELECT query (clause-based parameters, auto-joins and escapes table names).

**VBA Usage**
```vb
SqlQuery(selectClause, fromClause, [whereClause], [orderByClause], [filePath], outOk) As Variant()
```

| Parameter | Type | Description |
|------|------|------|
| selectClause | String | SELECT clause (e.g. "*" or "Name, Age") |
| fromClause | String | Table name (auto-adds [$] and bracket escaping) |
| whereClause | String | Optional. WHERE condition |
| orderByClause | String | Optional. ORDER BY clause |
| filePath | String | Optional. Workbook path |
| outOk | Boolean | output |

```vb
result = SqlQuery("*", "Sheet1", "Age > 30", "Name ASC", , ok)
```

#### CloseSqlCache

Release module-level cached ADODB connection. Recommended to call in Workbook_BeforeClose.

**VBA Usage**
```vb
CloseSqlCache
```

#### SqlGetConnection

Get ADODB database connection (module-level Static cache, reuse connection for same file). Prefers ACE (xlsx/xlsm), falls back to Jet (xls). **VBA-only**.

**VBA Usage**
```vb
SqlGetConnection([filePath], outOk) As Object
```

#### SqlRangeQuery

Query a Range directly, no saved workbook needed. Uses ADODB Recordset in memory. Supports SELECT * and WHERE filter.

**VBA Usage**
```vb
SqlRangeQuery(sql, rng, [tableAlias], outOk) As Variant()
```

| Parameter | Type | Description |
|------|------|------|
| sql | String | SQL statement (SELECT * FROM ... WHERE ...) |
| rng | Range | Data range (first row is column names) |
| tableAlias | String | Optional. Table alias, default "data" |
| outOk | Boolean | output |

```vb
result = SqlRangeQuery("SELECT * FROM data WHERE Score > 90", Range("A1:C100"), "data", ok)
```

### Recipe 4.2 -- Multi-Table Join

<a id="recipe-multi-table-join"></a>

**Scenario**: Join sales table with regions table, query the region name for each sales record.

**Input** (worksheet: `Sales`)
|   | A | B |
|---|---|---|
| 1 | Product | RegionID |
| 2 | Apple | 1 |
| 3 | Banana | 2 |

(worksheet: `Regions`)
|   | A | B |
|---|---|---|
| 1 | ID | RegionName |
| 2 | 1 | Beijing |
| 3 | 2 | Shanghai |

`=UDF_SQL_JOIN("Sales", "Regions", "t1.RegionID = t2.ID", "INNER", "t1.Product, t2.RegionName")` →

**Output**
|   | A | B |
|---|---|---|
| 1 | Product | RegionName |
| 2 | Apple | Beijing |
| 3 | Banana | Shanghai |

#### SqlJoin

Two-table JOIN query.

**VBA Usage**
```vb
SqlJoin(table1, table2, joinOn, [joinType], [selectCols], [filePath], outOk) As Variant()
```

| Parameter | Type | Description |
|------|------|------|
| table1 | String | Left table name |
| table2 | String | Right table name |
| joinOn | String | JOIN condition (e.g. "t1.ID = t2.ID") |
| joinType | String | Optional. "INNER" (default), "LEFT", "RIGHT" |
| selectCols | String | Optional. Output columns, default "*" |
| filePath | String | Optional. Workbook path |
| outOk | Boolean | output |

```vb
result = SqlJoin("Sales", "Regions", "t1.RegionID = t2.ID", "LEFT", , , ok)
```

**UDF Usage**
```
=UDF_SQL_JOIN(table1, table2, joinOn, [joinType], [selectCols], [filePath])
```

### Recipe 4.3 -- Group Aggregation

<a id="recipe-group-aggregation"></a>

**Scenario**: Group by region to count total sales amount and record count.

**Input** (worksheet: `Sales`)
|   | A | B |
|---|---|---|
| 1 | Region | Amount |
| 2 | Beijing | 1000 |
| 3 | Shanghai | 800 |
| 4 | Beijing | 1200 |
| 5 | Guangzhou | 600 |

`=UDF_SQL_GROUPBY("Sales", "Region", "SUM(Amount) AS Total, COUNT(*) AS Cnt")` →

**Output**
|   | A | B | C |
|---|---|---|---|
| 1 | Region | Total | Cnt |
| 2 | Beijing | 2200 | 2 |
| 3 | Shanghai | 800 | 1 |
| 4 | Guangzhou | 600 | 1 |

#### SqlGroupBy

Group aggregation query.

**VBA Usage**
```vb
SqlGroupBy(tableName, groupCols, aggExprs, [whereClause], [filePath], outOk) As Variant()
```

| Parameter | Type | Description |
|------|------|------|
| tableName | String | Table name |
| groupCols | String | Group columns (comma-separated) |
| aggExprs | String | Aggregation expressions (e.g. "SUM(Amount) AS Total") |
| whereClause | String | Optional. WHERE filter condition |
| filePath | String | Optional. Workbook path |
| outOk | Boolean | output |

```vb
result = SqlGroupBy("Sales", "Region", "SUM(Amount) AS Total, AVG(Amount) AS Avg", , , ok)
```

**UDF Usage**
```
=UDF_SQL_GROUPBY(tableName, groupCols, aggExprs, [whereClause], [filePath])
```

#### SqlListSheets

List all worksheet names in the workbook (with $ suffix).

**VBA Usage**
```vb
SqlListSheets([filePath], outOk) As Variant()
```

**UDF Usage**
```
=UDF_SQL_LIST_SHEETS([filePath])
```

#### SqlListColumns

List column names and their ordinals (ORDINAL_POSITION) for the specified table.

**VBA Usage**
```vb
SqlListColumns(tableName, [filePath], outOk) As Variant()
```

| Parameter | Type | Description |
|------|------|------|
| tableName | String | Table name (auto-adds $ suffix and brackets) |
| filePath | String | Optional. Workbook path |

**UDF Usage**
```
=UDF_SQL_LIST_COLUMNS(tableName, [filePath])
```

#### SqlListTables

List all available data sources in the workbook (including worksheets and workbook-level named ranges).

**VBA Usage**
```vb
SqlListTables([filePath], outOk) As Variant()
```

**UDF Usage**
```
=UDF_SQL_LIST_TABLES([filePath])
```

---

> **Environment Requirements**: SqlUtils depends on ADODB and ACE/Jet OLEDB providers. 64-bit Office requires Access Database Engine installation. If workbook is unsaved, use SqlRangeQuery to query a Range directly.
>
> **⚠️ SQL Injection Prevention**: ACE OLEDB Excel ISAM driver does not support parameterized queries. Use `SqlEscapeString()` to escape user-supplied cell values (`'` → `''`) before embedding them in WHERE/VALUES clauses: `=UDF_SQL_QUERY("SELECT * FROM [Sheet1$] WHERE Name = '" & SqlEscapeString(A1) & "'")`. Note: Error-valued cells (#N/A, #VALUE!, etc.) are silently converted to an empty string (to prevent `CStr` crashes) — upstream data containing errors may yield empty result sets; clean the data first.
>
> **Naming Convention**: Worksheet names in SQL require a `$` suffix and square brackets, e.g. `[Sheet1$]`. Column names with special characters are auto-cleaned by SqlRangeQuery.
---

## Chapter 5: LinearUtils — Matrix & Linear Algebra

Matrix operations, decomposition, solving, and worksheet array formulas. All output matrices are 1-based (consistent with Excel Range). **Module**: `LinearUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`MatrixRows`](#matrixrows) | `(A)` | Matrix row count | Long |
| [`MatrixCols`](#matrixcols) | `(A)` | Matrix column count | Long |
| [`MatrixFrobeniusNorm`](#matrixnorm) | `(A)` | Frobenius norm (scaled to prevent overflow) | Double |
| [`RangeToMatrix`](#rangetomatrix) | `(rng)` | Convert Range to Double(,) 1-based | Double(,) |
| [`MatrixToRange`](#matrixtorange) | `(mat, startCell)` | Write matrix to worksheet (Sub) | — |
| [`SelectionToArray2D`](#rangetomatrix) | `(rng)` | Convert selected range to Variant 2D array | Variant(,) |
| [`ArrayToRange`](#matrixtorange) | `(data, startCell, [asColumn])` | Write Variant array to Range (Sub) | — |
| [`MatrixTranspose`](#matrixtranspose) | `(A)` | Matrix transpose | Double(,) |
| [`IdentityMatrix`](#identitymatrix) | `(n)` | Generate nxn identity matrix | Double(,) |
| [`MatrixCopy`](#matrixcopy) | `(A)` | Deep copy matrix | Double(,) |
| [`MatrixGetColumn`](#matrixrows) | `(A, k)` | Extract k-th column (1-based) | Double() |
| [`MatrixSetColumn`](#matrixrows) | `(A, k, col)` | Set k-th column (Sub) | — |
| [`MatrixMultiply`](#matrixmultiply) | `(A, B, [blockSize])` | Block matrix multiplication (cache-friendly) | Double(,) |
| [`MatrixMultiplyNaive`](#matrixmultiply) | `(A, B)` | Naive triple-loop multiplication | Double(,) |
| [`MatrixScale`](#matrixscale) | `(A, scalar)` | scalar multiplication | Double(,) |
| [`MatrixAdd`](#matrixadd) | `(A, B)` | Matrix addition | Double(,) |
| [`MatrixSubtract`](#matrixsubtract) | `(A, B)` | Matrix subtraction A-B | Double(,) |
| [`MatrixHadamard`](#matrixhadamard) | `(A, B)` | Element-wise product (Hadamard) | Double(,) |
| [`MatrixPower`](#matrixpower) | `(A, n)` | Square matrix to power n (binary exponentiation, n>=0) | Double(,) |
| [`MatrixNorm`](#matrixnorm) | `(A, [normType])` | Matrix norm ("1"/"inf"/"fro") | Double |
| [`MatrixTrace`](#matrixtrace) | `(A)` | Trace (sum of diagonal) | Double |
| [`SVD`](#svd) | `(A, U, S, Vt, [tol])` | SVD One-Sided Jacobi (Sub) | ByRef |
| [`PseudoInverse`](#pseudoinverse) | `(A, [tolerance])` | Moore-Penrose pseudo-inverse | Double(,) |
| [`MatrixRank_Array`](#matrixrank_array) | `(A, [tolerance])` | Matrix rank based on SVD | Long |
| [`EigenSymmetric`](#eigensymmetric) | `(A, V, D, [tol])` | Symmetric eigendecomposition Jacobi (Sub) | ByRef |
| [`QRDecomposition`](#qrdecomposition) | `(A, Q, R, [economy])` | QR decomposition Householder (Sub) | ByRef |
| [`LUDecomposition`](#ludecomposition) | `(A, L, U, P, [swapCount])` | LU decomposition (Doolittle, partial pivoting) (Sub) | ByRef |
| [`CholeskyDecomposition`](#choleskydecomposition) | `(A, L)` | Cholesky decomposition A=LL^T (Sub) | ByRef |
| [`MatrixDeterminant`](#matrixdeterminant) | `(A)` | Determinant (via LU decomposition) | Double |
| [`SolveLinearSystem`](#solvelinearsystem) | `(A, b, [tolerance])` | Solve Ax=b (SVD pseudo-inverse) | Double() |
| [`MatrixConditionNumber`](#matrixconditionnumber) | `(A)` | Condition number max(S)/min(S) | Double |
| [`VectorDot`](#vectordot) | `(a, b)` | Vector dot product (Kahan compensation) | Double |
| [`VectorNorm`](#vectornorm) | `(v, [normType])` | Vector norm ("2"/"1"/"inf") | Double |
| [`VectorCross`](#vectorcross) | `(a, b)` | 3D vector cross product | Double(3) |
| [`PolyFit`](#polyfit) | `(rngX, rngY, [degree])` | Least squares Polynomial Fitting (QR method) | Variant(,1) |

| Recipe | Functions Used |
|--------|---------------|
| [SVD Singular Value Decomposition](#recipe-svd) | `UDF_LINALG_SVD_SVALS`, `UDF_LINALG_SVD_U`, `UDF_LINALG_SVD_VT` |
| [Linear System Solving](#recipe-linear-system) | `UDF_LINALG_SOLVE` |
| [Polynomial Fitting](#recipe-polyfit) | `UDF_LINALG_POLYFIT` |

<a id="recipe-svd"></a>
### Recipe 5.1 — SVD Singular Value Decomposition

**Scenario**: Perform SVD on a matrix, obtain singular values and left/right singular vectors, for dimensionality reduction, pseudo-inverse computation, or matrix compression.

**Input**
|   | A | B | C |
|---|---|---|---|
| 1 | 1 | 0 | 0 |
| 2 | 0 | 2 | 0 |
| 3 | 0 | 0 | 0 |

Select 3x1 range and enter array formula `=UDF_LINALG_SVD_SVALS(A1:C3)` → `{2; 1; 0}`

Select 3x3 range and enter `=UDF_LINALG_SVD_U(A1:C3)` → left singular vector matrix

`=UDF_LINALG_PINV(A1:C3, 1E-10)` → pseudo-inverse (3x3)

#### SVD

Perform Singular Value Decomposition A = U * S * V^T on any mxn matrix. Uses One-Sided Jacobi algorithm, numerically stable. Returns U (mxn), S (nxn diagonal matrix), V^T (nxn).

**VBA Usage**
```vb
SVD(A, U, S, Vt, [tol], [maxSweeps])
```

| Parameter | Type | Description |
|------|------|------|
| A | Double(,) | input matrix |
| U | Double(,) | output: left singular vector matrix (ByRef) |
| S | Double(,) | output: Singular value diagonal matrix (ByRef) |
| Vt | Double(,) | output: right singularvectortranspose (ByRef) |
| tol | Double | Optional. convergence tolerance |
| maxSweeps | Long | Optional. max scans |

```vb
Dim A() As Double, U() As Double, S() As Double, Vt() As Double
A = RangeToMatrix(Sheet1.Range("A1:C3"))
SVD A, U, S, Vt
MatrixToRange U, Sheet1.Range("E1")
' When m < n (wide matrix), automatically transposes to compute A^T then swaps U/Vt
```

> **What SVD gives you**: A = U * S * V^T. **S** contains singular values (diagonal, descending) — the "importance" of each component. The number of significant singular values equals the effective matrix rank. **U** columns are left singular vectors, **V^T** rows are right singular vectors. Common uses: dimensionality reduction (keep top k singular values), pseudo-inverse (`PseudoInverse`), and PCA-like analysis.

**UDF Usage**
```
=UDF_LINALG_SVD_U(rng, [tol], [maxSweeps])       → U matrix (left singular vector)
=UDF_LINALG_SVD_S(rng, [tol], [maxSweeps])       → S diagonal matrix
=UDF_LINALG_SVD_VT(rng, [tol], [maxSweeps])      → V^T matrix (right singular vector transpose)
=UDF_LINALG_SVD_SVALS(rng, [tol], [maxSweeps])   → Singular value column vector
```

| Parameter | Type | Description |
|------|------|------|
| rng | Range | Numeric matrix range |
| tol | Double | Optional. convergence tolerance |
| maxSweeps | Long | Optional. max scans |

#### PseudoInverse

Moore-Penrose pseudo-inverse based on SVD. Singular values below tolerance are truncated to zero. Negative tolerance auto-uses max(S) * eps * max(m,n).

**VBA Usage**
```vb
PseudoInverse(A, [tolerance]) As Double()
```

| Parameter | Type | Description |
|------|------|------|
| A | Double(,) | input matrix |
| tolerance | Double | Optional. Singular value truncation tolerance (-1=auto) |

```vb
Ap = PseudoInverse(A)
```

**UDF Usage**
```
=UDF_LINALG_PINV(rng, [tolerance])
```

#### EigenSymmetric

Symmetric matrix eigendecomposition A = V * D * V^T. Uses Jacobi rotation algorithm. Automatically checks symmetry; eigenvalues sorted descending.

**VBA Usage**
```vb
EigenSymmetric(A, V, D, [tol], [maxSweeps])
```

| Parameter | Type | Description |
|------|------|------|
| A | Double(,) | Symmetric square matrix |
| V | Double(,) | output: Eigenvector matrix (ByRef) |
| D | Double(,) | output: Eigenvalue diagonal matrix (ByRef) |

**UDF Usage**
```
=UDF_LINALG_EIGVAL(rng, [tol], [maxSweeps])   ' Eigenvalues
=UDF_LINALG_EIGVEC(rng, [tol], [maxSweeps])   ' Eigenvectors
```

<a id="recipe-linear-system"></a>
### Recipe 5.2 — Linear System Solving

**Scenario**: Solve linear system Ax = b. Supports exact solution for square matrices, least squares for overdetermined, and minimum norm for underdetermined systems.

**Input**
|   | A | B | C | D |
|---|---|---|---|---|
| 1 | 2 | 1 | -1 | 8 |
| 2 | -3 | -1 | 2 | -11 |
| 3 | -2 | 1 | 2 | -3 |

`=UDF_LINALG_SOLVE(A1:C3, D1:D3)` → `{2; 3; -1}` (solution x=2, y=3, z=-1)

#### SolveLinearSystem

Solve Ax = b using SVD pseudo-inverse. Non-singular square matrices yield exact solutions; overdetermined (m>n) yields least squares; underdetermined (m<n) yields minimum norm solution.

**VBA Usage**
```vb
SolveLinearSystem(A, b, [tolerance]) As Double()
```

| Parameter | Type | Description |
|------|------|------|
| A | Double(,) | mxn coefficient matrix |
| b | Double() | Right-hand side vector (length m) |
| tolerance | Double | Optional. SVD tolerance |

```vb
x = SolveLinearSystem(A, b)
' x(1)=2, x(2)=3, x(3)=-1
```

**UDF Usage**
```
=UDF_LINALG_SOLVE(rngA, rngB)
```

| Parameter | Type | Description |
|------|------|------|
| rngA | Range | Coefficient matrix A |
| rngB | Range | Right-hand side vector b (single column/row) |

#### MatrixDeterminant

Compute determinant of square matrix via LU decomposition. Sign determined by permutation count.

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

Condition number = max singular value / min singular value. Large values indicate ill-conditioned matrix. Returns MAX_DOUBLE when singular.

**VBA Usage**
```vb
MatrixConditionNumber(A, [tol], [maxSweeps]) As Double
```

<a id="recipe-polyfit"></a>
### Recipe 5.3 — Polynomial Fitting

**Scenario**: Given x, y data points, fit a polynomial using least squares, returns coefficients (high to low degree).

**Input**
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

Solve least squares using QR decomposition, avoiding the condition-squaring problem of normal equations. degree=1 returns [a,b] for y=ax+b; degree=2 returns [a,b,c] for y=ax^2+bx+c.

**VBA Usage**
```vb
PolyFit(rngX, rngY, [degree]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| rngX | Range/Array | Independent variable data (1D) |
| rngY | Range/Array | Dependent variable data (1D, same length as X) |
| degree | Long | Optional. polynomial degree (default 1) |

```vb
result = PolyFit(Range("A1:A5"), Range("B1:B5"), 2)
' result is 2D column vector, high to low degree
```

**UDF Usage**
```
=UDF_LINALG_POLYFIT(rngX, rngY, [degree])
```

| Parameter | Type | Description |
|------|------|------|
| rngX | Range | X Data range |
| rngY | Range | Y Data range (same row count as X) |
| degree | Long | Optional. polynomial degree (default 1) |

<a id="matrixrows"></a>
<a id="matrixcols"></a>
#### MatrixRows / MatrixCols

Matrix dimension queries.

```vb
m = MatrixRows(A)   ' Row count
n = MatrixCols(A)   ' Column count
```

<a id="rangetomatrix"></a>
<a id="matrixtorange"></a>
#### RangeToMatrix / MatrixToRange

Convert between Range and Double matrix. RangeToMatrix requires all cells to be numeric values.

```vb
A = RangeToMatrix(Sheet1.Range("A1:C3"))    ' Range -> Double(,)
MatrixToRange U, Sheet1.Range("E1")          ' Double(,) -> Range
```

<a id="matrixtranspose"></a>
<a id="matrixtranspose"></a>
<a id="identitymatrix"></a>
<a id="matrixcopy"></a>
#### MatrixTranspose / IdentityMatrix / MatrixCopy

Basic matrix operations.

```vb
At = MatrixTranspose(A)        ' A^T
I = IdentityMatrix(3)          ' 3x3 identity matrix
Acopy = MatrixCopy(A)          ' deep copy
```

<a id="matrixadd"></a>
<a id="matrixsubtract"></a>
<a id="matrixmultiply"></a>
<a id="matrixhadamard"></a>
#### MatrixAdd / MatrixSubtract / MatrixMultiply / MatrixHadamard

Matrix arithmetic. MatrixMultiply uses blocked algorithm for cache optimization. Throws error on dimension mismatch.

```vb
C = MatrixAdd(A, B)             ' A + B
D = MatrixSubtract(A, B)        ' A - B
E = MatrixMultiply(A, B, 64)    ' A * B (blocked)
F = MatrixHadamard(A, B)        ' A (Hadamard) B element-wise product
```

#### QRDecomposition

QR decomposition A = Q*R based on Householder reflections. Sub procedure, returns Q and R via ByRef.

**VBA Usage**
```vb
QRDecomposition A, Q, R, True      ' Economy mode: Q(mxk), R(kxk)
```

| Parameter | Type | Description |
|------|------|------|
| A | Double(,) | input matrix |
| Q | Double(,) | output: Orthogonal matrix (ByRef) |
| R | Double(,) | output: upper triangular matrix (ByRef) |
| economy | Boolean | Optional. Economy mode (defaults to True) |

**UDF Usage**
```
=UDF_LINALG_QR_Q(rng, [economy])   → Q matrix
=UDF_LINALG_QR_R(rng, [economy])   → R matrix
```

| Parameter | Type | Description |
|------|------|------|
| rng | Range | Numeric matrix range |
| economy | Boolean | Optional. Economy mode (defaults to False) |

#### QRDecompositionPiv

Column-pivoted QR decomposition (Businger-Golub) A*P = Q*R. Column pivoting improves numerical stability for rank-deficient detection, least squares, and subset selection.

**VBA Usage**
```vb
QRDecompositionPiv A, Q, R, perm, True   ' Economy mode + column pivoting
```

| Parameter | Type | Description |
|------|------|------|
| A | Double(,) | input matrix |
| Q | Double(,) | output: Orthogonal matrix (ByRef) |
| R | Double(,) | output: upper triangular matrix (ByRef) |
| perm | Long() | output: column permutation vector (ByRef), perm(k)=original column index |
| economy | Boolean | Optional. Economy mode (defaults to False) |

#### CholeskyDecomposition

Cholesky decomposition A = L*L^T for symmetric positive definite matrices. Sub procedure, returns L (lower triangular) via ByRef.

**VBA Usage**
```vb
CholeskyDecomposition A, L         ' A = L*L^T (lower triangular)
```

| Parameter | Type | Description |
|------|------|------|
| A | Double(,) | input matrix (required: symmetric positive definite) |
| L | Double(,) | output: lower triangular matrix (ByRef) |

**UDF Usage**
```
=UDF_LINALG_CHOLESKY(rng)          → L matrix (lower triangular)
```

| Parameter | Type | Description |
|------|------|------|
| rng | Range | Symmetric positive definite matrix range |

#### LUDecomposition

LU decomposition P*A = L*U using Doolittle method with partial pivoting. Sub procedure, returns L, U, P and row swap count via ByRef.

**VBA Usage**
```vb
LUDecomposition A, L, U, P, swaps  ' P*A = L*U
```

| Parameter | Type | Description |
|------|------|------|
| A | Double(,) | Input square matrix |
| L | Double(,) | Output: Lower triangular matrix (ByRef) |
| U | Double(,) | Output: Upper triangular matrix (ByRef) |
| P | Long(,) | Output: Permutation matrix (ByRef) |
| swaps | Long | Output: Row swap count (ByRef) |

<a id="vectordot"></a>
<a id="vectornorm"></a>
<a id="vectorcross"></a>
#### VectorDot / VectorNorm / VectorCross

Vector operations. Dot product uses Kahan compensated summation; cross product supports 3D vectors only.

```vb
dot = VectorDot(v1, v2)            ' Dot product (Kahan compensation)
nrm = VectorNorm(v, "2")           ' L2 norm (scaled)
cross = VectorCross(v1, v2)        ' Cross product -> {x, y, z}
```

#### MatrixRank_Array

Matrix rank based on SVD singular value tolerance.

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

Matrix scalar operations. MatrixNorm supports "1" (column sum norm), "inf" (row sum norm), "fro" (Frobenius). MatrixPower uses binary exponentiation.

```vb
tr = MatrixTrace(A)                ' Trace
n1 = MatrixNorm(A, "1")            ' 1-norm
B = MatrixScale(A, 3.5)            ' scalar multiplication
Ap = MatrixPower(A, 3)             ' A^3 (binary exponentiation)
```

---

## Chapter 6: StatsUtils — Descriptive Statistics, Inference & Distributions

Central tendency, dispersion, shape, ranking, correlation, data transformation, and hypothesis testing. **Module**: `StatsUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`Mean`](#mean) | `(data, [colIndex])` | Arithmetic mean (Kahan compensation) | Variant (Double) |
| [`WeightedMean`](#mean) | `(values, weights)` | Weighted arithmetic mean | Variant (Double) |
| [`Max`](#max) | `(data, [colIndex])` | Maximum value (VBA-only) | Variant (Double) |
| [`Median`](#median) | `(data, [colIndex])` | Median | Variant (Double) |
| [`Min`](#min) | `(data, [colIndex])` | Minimum value (VBA-only) | Variant (Double) |
| [`MinMax`](#minmax) | `(data, [outMin], [outMax])` | Return min and max simultaneously | Variant Array |
| [`Mode`](#mode) | `(data, [colIndex])` | Mode (Returns #N/A when no unique mode) | Variant (Double) |
| [`GeometricMean`](#geometricmean) | `(data, [colIndex])` | Geometric mean (log-space for overflow safety) | Variant (Double) |
| [`HarmonicMean`](#harmonicmean) | `(data, [colIndex])` | Harmonic mean | Variant (Double) |
| [`TrimMean`](#trimmean) | `(data, [trimPct])` | Trimmed mean (trims trimPct/2 from both ends) | Variant (Double) |
| [`RootMeanSquare`](#rootmeansquare) | `(data, [colIndex])` | Root mean square RMS | Variant (Double) |
| [`MeanAbsDev`](#meanabsdev) | `(data, [colIndex])` | Mean absolute deviation | Variant (Double) |
| [`StdDev`](#stddev) | `(data, [colIndex])` | Sample standard deviation (n-1) | Variant (Double) |
| [`StdDevP`](#stddevp) | `(data, [colIndex])` | Population standard deviation (n) | Variant (Double) |
| [`Variance`](#variance) | `(data, [colIndex])` | Sample variance (n-1, Kahan compensation) | Variant (Double) |
| [`VarianceP`](#variancep) | `(data, [colIndex])` | Population variance (n) | Variant (Double) |
| [`Percentile`](#percentile) | `(data, k, [colIndex])` | Percentile (linear interpolation, 0<=k<=1) | Variant (Double) |
| [`IQR`](#iqr) | `(data, [colIndex])` | Interquartile range Q3-Q1 | Variant (Double) |
| [`Skewness`](#skewness) | `(data, [colIndex])` | Sample skewness (Fisher-Pearson) | Variant (Double) |
| [`Kurtosis`](#kurtosis) | `(data, [colIndex])` | Excess kurtosis | Variant (Double) |
| [`Rank`](#rank) | `(data, value, [ascending])` | modified competition ranking (ties counted as <=) | Variant (Long) |
| [`RankEq`](#rankeq) | `(data, value, [ascending])` | Consistent with Excel RANK.EQ | Variant (Long) |
| [`RankAvg`](#rankavg) | `(data, value, [ascending])` | Consistent with Excel RANK.AVG | Variant (Double) |
| [`PercentRank`](#percentrank) | `(data, value, [ascending])` | Percentile rank (0-1) | Variant (Double) |
| [`Covariance`](#covariance) | `(dataX, dataY)` | Sample covariance (pairwise deletion) | Variant (Double) |
| [`Correlation`](#correlation) | `(dataX, dataY)` | Pearson correlation coefficient | Variant (Double) |
| [`RSquare`](#rsquare) | `(actual, predicted)` | R^2 coefficient of determination | Variant (Double) |
| [`CorrelationMatrix`](#correlationmatrix) | `(data, [hasHeader])` | Numeric column correlation coefficient matrix (with labels) | Variant(,) |
| [`ZScore`](#zscore) | `(data, [value])` | Z-score normalization (single value if value given) | Variant |
| [`Normalize`](#normalize) | `(data, [colIndex])` | Min-Max normalization to [0, 1] | Double() |
| [`LinInterp`](#lininterp) | `(x, xs, ys)` | Linear interpolation (returns boundary values if out of range) | Variant (Double) |
| [`Winsorize`](#winsorize) | `(data, [pct])` | Winsorizing (both ends pct/2) | Double() |
| [`MovingAverage`](#movingaverage) | `(data, window)` | Simple moving average | Double() |
| [`Binning`](#binning) | `(data, nBins)` | equal-width binning **VBA-only** | Dictionary |
| [`ZTest`](#ztest) | `(data, mu0, sigma)` | One-sample Z-test (two-tailed p-value) | Variant (Double) |
| [`TTest`](#ttest) | `(data1, data2, [testType])` | Two-sample t-test (1=paired/2=equal variance/3=Welch) | Variant (Double) |
| [`StandardError`](#standarderror) | `(data, [colIndex])` | Standard error of the mean SE = s/sqrt(n) | Variant (Double) |
| [`ConfidenceInterval`](#confidenceinterval) | `(data, [alpha])` | t-distribution confidence interval **VBA-only** | Dictionary |
| [`GammaLn`](#gammaln) | `(x)` | Log Gamma function (Lanczos) | Double |
| [`BetaReg`](#betareg) | `(x, a, b)` | Regularized incomplete Beta function I_x(a,b) | Double |
| [`TDistCDF`](#tdistcdf) | `(tVal, df)` | t-distribution upper-tail probability P(T>|t|) | Double |
| [`TDist2T`](#tdist2t) | `(tStat, df)` | t-distribution two-tailed P-value 2*P(T>|t|) | Double |
| [`FDistRT`](#fdistrt) | `(fStat, df1, df2)` | F-distribution right-tail probability P(F>f) | Double |
| [`TInv2T`](#tinv2t) | `(alpha, df)` | t-distribution two-tailed critical value (binary search) | Double |

| Recipe | Functions Used |
|--------|---------------|
| [Descriptive Statistics](#recipe-descriptive-stats) | `UDF_STAT_MEAN`, `UDF_STAT_STDEV`, `UDF_STAT_MEDIAN`, `UDF_STAT_IQR` |
| [Hypothesis Testing](#recipe-hypothesis-test) | `UDF_STAT_ZTEST`, `UDF_STAT_TTEST` |
| [Correlation Analysis](#recipe-correlation) | `UDF_STAT_CORREL`, `UDF_STAT_COV` |

#### Min

#### Max

Return the minimum and maximum value. Returns Empty when no numeric values present. **VBA-only**.

**VBA Usage**

```vb
Min(data, [colIndex]) As Variant
Max(data, [colIndex]) As Variant
```

| Parameter | Type | Description |
|-----------|------|-------------|
| data | Variant | 1D or 2D array |
| colIndex | Long | Optional. 1-based column index for 2D arrays (default 1) |

```vb
Min(Array(3, 1, 5, 2))   ' → 1
Max(Array(3, 1, 5, 2))   ' → 5
```

<a id="recipe-descriptive-stats"></a>
### Recipe 6.1 — Descriptive Statistics

**Scenario**: For a set of experimental data, compute mean, standard deviation, median, and interquartile range to quickly understand data distribution characteristics.

**Input**
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

`=UDF_STAT_STDEV(A1:A7)` → `12.13` (Sample standard deviation, affected by outlier 45.0)

`=UDF_STAT_IQR(A1:A7)` → `1.05` (Interquartile range, robust to outliers)

`=UDF_STAT_SKEW(A1:A7)` → `2.14` (Severely right-skewed)

`=UDF_STAT_PERCENTILE(A1:A7, 0.9)` → `45.0` (90th percentile)

> **How to read these results**: Mean (17.54) is pulled up by the outlier 45.0 — it overstates the typical value. Median (13.1) is more representative here because it ignores the outlier. The gap between Mean and Median is a red flag for skewed data. IQR (1.05) tells you the middle 50% of values span only 1.05 units — the data is tightly clustered except for the outlier. Always check Skewness: >1 or <-1 indicates severe skew; near 0 is symmetric. Use StdDev (sample) when working with a sample of a larger population; use StdDevP when you have data for the entire population.

<a id="mean"></a>
<a id="median"></a>
<a id="mode"></a>
<a id="minmax"></a>
#### Mean / Median / Mode / MinMax

Central tendency measures. Mean uses Kahan compensated summation. Mode returns #N/A when no duplicate values. MinMax returns minimum and maximum simultaneously, suitable for array formulas.

**VBA Usage**
```vb
Mean(data, [colIndex]) As Variant
Median(data, [colIndex]) As Variant
Mode(data, [colIndex]) As Variant
MinMax(data, [outMin], [outMax], [colIndex]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| data | Range/Array | Numeric data |
| colIndex | Long | Optional. Column index for 2D arrays (default 1) |

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

Dispersion measures. Sample version divides by n-1, population version divides by n. All use Kahan compensated summation.

**VBA Usage**
```vb
StdDev(data, [colIndex]) As Variant      ' Sample standard deviation
StdDevP(data, [colIndex]) As Variant     ' Population standard deviation
Variance(data, [colIndex]) As Variant    ' Sample variance
VarianceP(data, [colIndex]) As Variant   ' Population variance
```

```vb
s = StdDev(Range("A1:A100"))     ' Sample standard deviation
s2 = Variance(Range("A1:A100"))  ' Sample variance
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

Other mean indicators. GeometricMean suitable for growth rate scenarios; HarmonicMean suitable for speed/rate scenarios; TrimMean removes extreme values; RMS used for signal strength; MAD more robust to outliers.

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
t = TrimMean(Range("A1:A100"), 0.1)          ' Trims 5% from each end
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

Distribution shape and percentiles. Percentile uses linear interpolation; Skewness >0 right-skewed, <0 left-skewed; Kurtosis >0 heavy-tailed.

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
### Recipe 6.2 — Hypothesis Testing

**Scenario**: Given population standard deviation sigma=5, test if sample mean significantly deviates from mu0=100; or compare mean differences between two groups.

**Input**
|   | A |
|---|---|
| 1 | 98 |
| 2 | 102 |
| 3 | 95 |
| 4 | 100 |
| 5 | 97 |
| 6 | 103 |
| 7 | 99 |

Z test: sigma=5, mu0=100

`=UDF_STAT_ZTEST(A1:A7, 100, 5)` → `0.596` (p > 0.05, do not reject H0)

Paired t-test, comparing group A vs group B:

|   | A | B |
|---|---|---|
| 1 | 85 | 88 |
| 2 | 90 | 92 |
| 3 | 78 | 82 |
| 4 | 92 | 95 |
| 5 | 88 | 91 |

`=UDF_STAT_TTEST(A1:A5, B1:B5, 1)` → `0.002` (p < 0.01, significant difference)

> **How to read p-values**: The p-value is the probability of seeing your data (or more extreme) if there were truly no effect. Lower p = stronger evidence against the null hypothesis.
> - **p < 0.01**: Strong evidence — the difference is statistically significant.
> - **p < 0.05**: Moderate evidence — commonly used threshold for significance.
> - **p > 0.05**: Weak evidence — cannot reject the null hypothesis; the observed difference may be due to chance.
>
> **Which test to use**: Use `ZTest` when you know the population standard deviation (large sample sizes). Use `TTest` for small samples or when sigma is unknown. For TTest, `testType=1` for paired data (same subjects before/after), `testType=2` for equal-variance groups, `testType=3` (Welch) for unequal-variance groups.

#### ZTest

One-sample Z test, two-tailed p value = 2 * (1 - Phi(|z|)). Uses Abramowitz-Stegun normal CDF approximation.

**VBA Usage**
```vb
ZTest(data, mu0, sigma, [colIndex]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| data | Range/Array | sampledata |
| mu0 | Double | Null hypothesis mean |
| sigma | Double | Known population standard deviation |
| colIndex | Long | Optional. Column index for 2D arrays |

```vb
p = ZTest(Range("A1:A30"), 100#, 5#)
' p < 0.05 -> reject H0
```

**UDF Usage**
```
=UDF_STAT_ZTEST(data, mu0, sigma, [colIndex])
```

#### TTest

Two-sample t-test. testType: 1=paired, 2=equal variance (default), 3=Welch. p value based on Pure VBA t distribution CDF (via incomplete Beta function).

**VBA Usage**
```vb
TTest(data1, data2, [testType], [colIdx1], [colIdx2]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| data1 | Range/Array | First group data |
| data2 | Range/Array | Second group data |
| testType | Long | Optional. 1=paired, 2=equal variance (default), 3=Welch |

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

Standard error SE = s / sqrt(n). ConfidenceInterval returns t-distribution confidence interval dictionary (VBA-only), with lower/upper/mean/se keys.

**VBA Usage**
```vb
StandardError(data, [colIndex]) As Variant
ConfidenceInterval(data, [alpha], [colIndex]) As Object  ' VBA-only
```

```vb
se = StandardError(Range("A1:A30"))
Set ci = ConfidenceInterval(Range("A1:A30"), 0.05)
' ci("lower") -> lower bound, ci("upper") -> upper bound
```

**UDF Usage**
```
=UDF_STAT_SE(data, [colIndex])
```

<a id="recipe-correlation"></a>
### Recipe 6.3 — Correlation Analysis

**Scenario**: Analyze linear correlation strength and direction between two numeric data sets.

**Input**
|   | A | B |
|---|---|---|
| 1 | 1 | 2.1 |
| 2 | 2 | 4.0 |
| 3 | 3 | 6.3 |
| 4 | 4 | 7.9 |
| 5 | 5 | 10.2 |

`=UDF_STAT_CORREL(A1:A5, B1:B5)` → `0.999` (Very strong positive correlation)

`=UDF_STAT_COV(A1:A5, B1:B5)` → `4.03` (Positive covariance)

`=UDF_STAT_R2(A1:A5, B1:B5)` → `0.998` (Good fit)

> **How to read correlation**: The Pearson correlation coefficient `r` ranges from -1 to +1. **r > 0.8** or **r < -0.8** indicates strong correlation; **0.5 < |r| < 0.8** is moderate; **|r| < 0.3** is weak. R² (coefficient of determination) tells you what fraction of Y's variance is explained by X — here R²=0.998 means 99.8% of variation in B is explained by A. Covariance indicates the direction of the relationship but its magnitude depends on the units; use Correlation for a standardized measure.

<a id="correlation"></a>
<a id="covariance"></a>
<a id="rsquare"></a>
#### Correlation / Covariance / RSquare

Correlation analysis. Correlation is the Pearson correlation coefficient r; Covariance is sample covariance; RSquare is R^2 = 1 - SSE/SST. All use pairwise deletion to handle missing values.

**VBA Usage**
```vb
Covariance(dataX, dataY, [colX], [colY]) As Variant
Correlation(dataX, dataY, [colX], [colY]) As Variant
RSquare(actual, predicted, [colActual], [colPredicted]) As Variant
```

| Parameter | Type | Description |
|------|------|------|
| dataX | Range/Array | First variable |
| dataY | Range/Array | Second variable |
| colX/colY | Long | Optional. Column index for 2D arrays |

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

Compute Pearson correlation coefficient matrix between multiple numeric data columns. Uses pairwise complete observations, Kahan compensated summation. Output is 2D Variant array with row/column labels.

**VBA Usage**
```vb
CorrelationMatrix(data, [hasHeader]) As Variant()
```

```vb
Dim corr As Variant
corr = CorrelationMatrix(Range("A1").CurrentRegion)
' Output: matrix with "Correlation" label
```

**UDF Usage** (via RegressUtils proxy)
```
=UDF_REGRESS_CORREL(data)
```

<a id="rank"></a>
<a id="rankeq"></a>
<a id="rankavg"></a>
<a id="percentrank"></a>
#### Rank / RankEq / RankAvg / PercentRank

Ranking functions. Rank is modified competition ranking; RankEq equivalent to Excel RANK.EQ; RankAvg equivalent to RANK.AVG (ties averaged); PercentRank returns (rank-1)/(n-1).

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

Data transformation tools. ZScore returns single Z score if a single value is given, otherwise returns full array; Normalize is Min-Max normalization; LinInterp performs linear interpolation on sorted xs, ys; Winsorize compresses extreme values; MovingAverage is simple moving average; Binning is equal-width binning (VBA-only, returns Dictionary).

**VBA Usage**
```vb
ZScore(data, [value], [colIndex]) As Variant
Normalize(data, [colIndex]) As Variant
LinInterp(x, xs, ys) As Variant
Winsorize(data, [pct], [colIndex]) As Variant
MovingAverage(data, window, [colIndex]) As Variant
Binning(data, nBins, [colIndex]) As Object        ' VBA-only
```

```vb
z = ZScore(Range("A1:A100"), 75#)
norm = Normalize(Range("A1:A100"))
y = LinInterp(2.5, Array(1,2,3), Array(10,20,30))  ' -> 25
w = Winsorize(Range("A1:A100"), 0.05)
ma = MovingAverage(Range("A1:A100"), 5)
Set bins = Binning(Range("A1:A100"), 10)
' bins("__internal_edges__") -> Boundary array
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

Log Gamma function ln(Γ(x)). Based on Lanczos approximation (g=7, 9-term coefficients), precision ~1E-12. Pure VBA implementation, no Excel dependency.

```vb
GammaLn(x) As Double
```

| Parameter | Type | Description |
|------|------|------|
| x | Double | Positive value, x > 0 |

```vb
gl = GammaLn(1#)     ' → 0      (ln(Γ(1)) = ln(1))
gl = GammaLn(5#)     ' → 3.178  (ln(24))
gl = GammaLn(0.5)    ' → 0.572  (ln(√π))
```

#### BetaReg

Regularized incomplete beta function I_x(a,b). Uses Modified Lentz continued fraction, precision ~1E-12. Pure VBA implementation.

```vb
BetaReg(x, a, b) As Double
```

| Parameter | Type | Description |
|------|------|------|
| x | Double | Integration upper limit, 0 ≤ x ≤ 1 |
| a | Double | Beta shape parameter 1, a > 0 |
| b | Double | Beta shape parameter 2, b > 0 |

```vb
betaI = BetaReg(0.5, 2#, 3#)    ' → ~0.6875
betaI = BetaReg(1#, 1#, 1#)     ' → 1.0    (Uniform distribution)
```

#### TDistCDF

t-distribution upper-tail probability P(T > |t|). Via BetaReg implementation. Pure VBA, replaces Excel TDIST.

```vb
TDistCDF(tVal, df) As Double
```

| Parameter | Type | Description |
|------|------|------|
| tVal | Double | t statistic, tVal >= 0 |
| df | Double | Degrees of freedom, df > 0 |

```vb
p = TDistCDF(2.5, 15)   ' P(T > 2.5), df=15
p = TDistCDF(1.96, 100) ' P(T > 1.96), df=100
```

#### TDist2T

t-distribution two-tailed P-value 2 × P(T > |t|). Used for two-tailed t-test.

```vb
TDist2T(tStat, df) As Double
```

| Parameter | Type | Description |
|------|------|------|
| tStat | Double | t statistic |
| df | Double | Degrees of freedom, df > 0 |

```vb
p2 = TDist2T(2.5, 15)   ' two-tailed P-value, df=15
p2 = TDist2T(3.0, 20)   ' two-tailed P-value, df=20
```

#### FDistRT

F-distribution right-tail probability P(F > f). Via BetaReg implementation. Pure VBA, replaces Excel FDIST.

```vb
FDistRT(fStat, df1, df2) As Double
```

| Parameter | Type | Description |
|------|------|------|
| fStat | Double | F statistic, fStat >= 0 |
| df1 | Double | Numerator df, df1 > 0 |
| df2 | Double | Denominator df, df2 > 0 |

```vb
pF = FDistRT(3.2, 3, 20)    ' P(F > 3.2), df=(3,20)
pF = FDistRT(5.0, 2, 30)    ' P(F > 5.0), df=(2,30)
```

#### TInv2T

t-distribution two-tailed critical value. Given significance level alpha, returns t critical value such that P(|T| > t) = alpha. Uses binary search, precision ~1E-12.

```vb
TInv2T(alpha, df) As Double
```

| Parameter | Type | Description |
|------|------|------|
| alpha | Double | Significance level, 0 < alpha < 1 |
| df | Double | Degrees of freedom, df > 0 |

```vb
tCrit = TInv2T(0.05, 20)    ' t critical value (alpha=0.05 two-tailed, df=20)
tCrit = TInv2T(0.01, 10)    ' t critical value (alpha=0.01 two-tailed, df=10)
```

---

## Chapter 7: RegressUtils — Regression, ANOVA & Factor Optimization

Multiple linear regression (with categorical/Boolean factor encoding), one-way ANOVA, interaction effects detection, factor sweep and grid optimization. **Module**: `RegressUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`FactorImportance`](#factorimportance) | `(data, factorCols, resultCol)` | Factor importance ranking (normalized coefficients) | Variant(,) |
| [`InteractionEffects`](#interactioneffects) | `(data, factorCols, resultCol)` | Pairwise interaction effects F test | Variant(,) |
| [`ANOVAOneWay`](#anovaoneway) | `(data, factorCol, resultCol)` | One-way ANOVA **VBA-only** | Dictionary |
| [`ANOVAOneWay_Fstat`](#anovaoneway) | `(data, factorCol, resultCol)` | One-way ANOVA F-statistic only | Double |
| [`LinearModelFit`](#linearmodelfit) | `(data, factorCols, resultCol)` | Multiple linear regression fitting **VBA-only** | Dictionary |
| [`LinearModelPredict`](#linearmodelpredict) | `(model, newData)` | Predict result from model | Double |
| [`FactorSweep`](#factorsweep) | `(model, factorIdx, from, to, steps)` | Single-factor What-If sweep | Variant(,) |
| [`OptimizeFactors`](#optimizefactors) | `(data, factorCols, resultCol, [goal], [topN])` | Grid search optimal factor combinations | Variant(,) |

| Recipe | Functions Used |
|--------|---------------|
| [Factor Importance Analysis](#recipe-factor-importance) | `UDF_REGRESS_IMPORTANCE` |
| [one-way ANOVA](#recipe-one-way-anova) | `UDF_REGRESS_ANOVA` |
| [Factor optimization](#recipe-factor-optimization) | `UDF_REGRESS_OPTIMIZE` |

<a id="recipe-factor-importance"></a>
### Recipe 7.1 — Factor Importance Analysis

**Scenario**: Multiple factors affect product quality metrics; need to quantify the relative importance of each factor and identify key drivers.

**Input**
|   | A | B | C | D |
|---|---|---|---|---|
| 1 | Temp | Pressure | Time | Yield |
| 2 | 100 | 2.5 | 30 | 85.2 |
| 3 | 110 | 2.8 | 35 | 88.1 |
| 4 | 120 | 3.0 | 40 | 91.3 |
| 5 | 130 | 3.2 | 45 | 94.5 |
| 6 | 140 | 3.5 | 50 | 97.8 |

Select the factor column name range (e.g. A1:C1) as the factorCols parameter:

`=UDF_REGRESS_IMPORTANCE(A1:D6, A1:C1, 4)` ->

| Rank | Factor | Normalized Coeff | Raw Coeff | Absolute Importance |
|------|------|-----------|---------|-----------|
| 1 | Temperature | 0.852 | 0.321 | 0.852 |
| 2 | Pressure | 0.456 | 15.23 | 0.456 |
| 3 | Time | 0.213 | 0.089 | 0.213 |

#### FactorImportance

Importance ranking based on normalized regression coefficients. Auto-detects factor types (numeric/Boolean/categorical), categorical factors auto-encoded as k-1 dummy variables then aggregated. Output is 2D array with headers: rank, factor name, normalized coefficient, raw coefficient, absolute importance.

> **How to read the output**:
> - **Rank**: 1 = most important factor. Sorted by absolute importance descending.
> - **Normalized coefficient**: Coefficients scaled to [-1, 1] for fair comparison across factors with different units. Larger absolute value = more impact.
> - **Raw coefficient**: The actual regression coefficient in original units. Temperature's raw coefficient of 0.321 means yield increases by 0.321 units per degree.
> - **Absolute importance**: Same as normalized coefficient in the current implementation. Use this column to rank factors.
>
> In the example: Temperature (0.852) is ~2× more important than Pressure (0.456) and ~4× more than Time (0.213).

**VBA Usage**
```vb
FactorImportance(data, factorCols, resultCol, [hasHeader]) As Variant()
```

| Parameter | Type | Description |
|------|------|------|
| data | Range/Variant(,) | Data range (first row defaults to header) |
| factorCols | Variant Array | factorcolumnindexarrays (1-based) |
| resultCol | Long | resultcolumnindex (1-based) |
| hasHeader | Boolean | Optional. Whether includes header (defaults to True) |

```vb
Dim importance As Variant
importance = FactorImportance(dataRange, Array(1,2,3), 4)
' outputincludes header, normalizedlarger coefficient -> factormore important
```

**UDF Usage**
```
=UDF_REGRESS_IMPORTANCE(data, factorCols, resultCol)
```

| Parameter | Type | Description |
|------|------|------|
| data | Range | Full data range (includes header) |
| factorCols | Range | Factor column name range (e.g. A1:C1) |
| resultCol | Long | Result column index (1-based) |

<a id="recipe-one-way-anova"></a>
### Recipe 7.2 — One-Way ANOVA

**Scenario**: Compare whether product strength has significant differences under three different process parameter settings.

**Input**
|   | A | B |
|---|---|---|
| 1 | Process | Strength |
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

One-way ANOVA, tests whether a categorical factor has significant effect on numeric results. Uses two-pass within-group sum of squares to avoid catastrophic cancellation. Returns Dictionary (VBA-only) with SSB/SSW/SST/MSB/MSW/F/p_value/eta_sq/significant/summary keys.

**VBA Usage**
```vb
ANOVAOneWay(data, factorCol, resultCol, [hasHeader]) As Object  ' VBA-only
```

| Parameter | Type | Description |
|------|------|------|
| data | Range/Variant(,) | Data range |
| factorCol | Long | Categorical factor column index (1-based) |
| resultCol | Long | Numeric result column index (1-based) |
| hasHeader | Boolean | Optional. Defaults to True |

`ANOVAOneWay_Fstat` is a convenience wrapper returning only the F-statistic as a Double (same parameters).

```vb
Set anova = ANOVAOneWay(dataRange, 1, 2)
f = ANOVAOneWay_Fstat(dataRange, 1, 2)   ' → 8.42
' anova("summary") -> "F(2,6) = 25.500, p = 0.0012, eta^2 = 0.895"
' anova("F") -> 25.5
' anova("p_value") -> 0.0012
' anova("significant") -> True
```

> **How to read the output**: The Dictionary returned by `ANOVAOneWay` contains these keys:
> 
> | Key | Meaning | How to interpret |
> |-----|---------|-----------------|
> | `SSB` / `SSW` / `SST` | Sum of Squares: Between / Within / Total | SSB >> SSW means groups differ |
> | `MSB` / `MSW` | Mean Squares | Used to compute F |
> | `F` | F-statistic = MSB/MSW | Larger F → stronger group differences |
> | `p_value` | Probability of this F by chance | **p < 0.05** → groups differ significantly |
> | `eta_sq` | Effect size (SSB/SST) | >0.14 large, >0.06 medium, <0.01 small |
> | `significant` | Boolean | Quick True/False at α=0.05 |
> | `summary` | Readable string | e.g. "F(2,6)=25.500, p=0.0012, eta²=0.895" |

**UDF Usage**
```
=UDF_REGRESS_ANOVA(data, factorCol, resultCol)
```

| Parameter | Type | Description |
|------|------|------|
| data | Range | Data range (includes header) |
| factorCol | Long | Categorical factor column index |
| resultCol | Long | Numeric result column index |

<a id="linearmodelfit"></a>
<a id="linearmodelpredict"></a>
#### LinearModelFit / LinearModelPredict

Multiple linear regression fitting and prediction. Builds design matrix with intercept, numeric factors, Boolean factors, and categorical dummy variables; solves coefficients using QR decomposition, computes standard errors via (X^T X)^(-1) pseudo-inverse. Returns Dictionary with coefficients/coef_names/r_squared/adj_r_squared/se/t_stats/p_values/f_stat/f_pvalue/fitted_values/residuals/factor_map/formula keys. LinearModelPredict predicts new data based on the model.

**VBA Usage**
```vb
LinearModelFit(data, factorCols, resultCol, [hasHeader]) As Object  ' VBA-only
LinearModelPredict(model, newData) As Double
```

| Parameter | Type | Description |
|------|------|------|
| data | Range/Variant(,) | Training data |
| factorCols | Variant | factorcolumnindexarrays |
| resultCol | Long | resultcolumnindex |
| newData | Range/Array | New factor values for prediction (in factorCols order) |

```vb
Set model = LinearModelFit(dataRange, Array(1,2,3), 4)
' model("r_squared") -> R^2
' model("formula") -> "Y ~ Temperature + Pressure + Time"
' model("coefficients") -> Coefficient array

pred = LinearModelPredict(model, Array(120, 3.0, 40))
' -> predictionvalue
```

> **How to read the output**: The Dictionary returned by `LinearModelFit` contains:
> 
> | Key | Meaning | How to interpret |
> |-----|---------|-----------------|
> | `coefficients` | Regression coefficients (β) | First value is intercept; remaining match `coef_names` order |
> | `coef_names` | Factor names | Shows which coefficient belongs to which factor |
> | `r_squared` | R²: fraction of variance explained | **>0.7** good fit; **>0.9** excellent; close to 1 = model explains nearly all variation |
> | `adj_r_squared` | Adjusted R² (penalizes extra factors) | Compare with `r_squared`; large gap = some factors may be irrelevant |
> | `se` | Standard errors of coefficients | Smaller SE = more precise estimate |
> | `t_stats` | t-statistics = coef / se | **|t| > 2** suggests the factor is meaningful |
> | `p_values` | Probability each coefficient is zero | **p < 0.05** → factor is statistically significant |
> | `f_stat` / `f_pvalue` | Overall model F-test | **p < 0.05** → at least one factor matters |
> | `fitted_values` | Model predictions for training data | Compare with actual values to assess fit |
> | `residuals` | Actual - predicted | Should be randomly scattered (no pattern) |
> | `factor_map` | Internal encoding map | Technical: how categorical factors were encoded |
> | `formula` | Human-readable model formula | e.g. "Y ~ Temperature + Pressure + Time" |

**UDF Usage**
```
=UDF_REGRESS_PREDICT(data, factorCols, resultCol, newVals)
```

#### InteractionEffects

Pairwise interaction effects detection. For each factor pair, builds a full model with interaction terms; uses F test to compare full model vs baseline model residual sum of squares to check if improvement is significant. Supports any combination of numeric/Boolean/categorical factors.

**VBA Usage**
```vb
InteractionEffects(data, factorCols, resultCol, [hasHeader]) As Variant()
```

```vb
Dim interact As Variant
interact = InteractionEffects(dataRange, Array(1,2,3,4), 5)
' Output includes header "factorA/factorB/interaction terms count/F value/p value/Significant"
' p < 0.05 -> Interaction effects significant
```

> **How to read the output**: Each row tests whether two factors interact (their combined effect differs from the sum of individual effects). **p < 0.05** means a significant interaction — the factors don't act independently. For example, Temperature×Pressure p=0.003 means the effect of raising temperature depends on the pressure level. A non-significant interaction (p>0.05) means the factors can be optimized independently.

**UDF Usage**
```
=UDF_REGRESS_INTERACT(data, factorCols, resultCol)
```

<a id="recipe-factor-optimization"></a>
### Recipe 7.3 — Factor Optimization

**Scenario**: Search for process parameter combinations that maximize yield under low temperature and high pressure conditions, while keeping the search space manageable.

**Input**: Same data as Recipe 7.1 (Temperature 100-140, Pressure 2.5-3.5, Time 30-50)

`=UDF_REGRESS_OPTIMIZE(A1:D6, A1:C1, 4, "max", 3, 10)` ->

| Temp | Pressure | Time | Predicted |
|------|------|------|-----------|
| 140 | 3.5 | 50 | 97.8 |
| 135.6 | 3.39 | 47.8 | 96.2 |
| 131.1 | 3.28 | 45.6 | 94.8 |

#### OptimizeFactors

Grid search for optimal factor combinations. Generates nSteps discrete points within observed range for numeric factors, enumerates {0,1} for Boolean factors, and all levels for categorical factors. Search space upper bound MAX_GRID_COMBOS=200k. goal="max" maximizes, "min" minimizes, or pass a numeric value for target approximation.

**VBA Usage**
```vb
OptimizeFactors(data, factorCols, resultCol, [goal], [topN], [nSteps], [hasHeader]) As Variant()
```

| Parameter | Type | Description |
|------|------|------|
| data | Range/Variant(,) | Training data |
| factorCols | Variant | factorcolumnindexarrays |
| resultCol | Long | Result column index |
| goal | Variant | Optional. "max" / "min" / numeric value (default "max") |
| topN | Long | Optional. Number of optimal combinations to return (default 10) |
| nSteps | Long | Optional. Discrete steps for numeric factors (default 10) |

```vb
Dim best As Variant
best = OptimizeFactors(dataRange, Array(1,2,3), 4, "max", 5, 10)
' Returns top 5 optimal combinations
best = OptimizeFactors(dataRange, Array(1,2), 3, CDbl(7.5), 3, 20)
' Returns 3 combinations closest to 7.5
```

> **How to read the output**: Returns a 2D array with columns: rank, predicted value, and one column per factor showing the optimal setting. Each row is a candidate combination sorted by desirability.
> - **goal="max"**: Rows sorted by predicted value descending (best first).
> - **goal="min"**: Rows sorted by predicted value ascending.
> - **goal=numeric**: Rows sorted by proximity to target value (closest first).
> - **topN**: Controls how many top results are returned. For multimodal surfaces, review multiple candidates (not just #1) — a slightly lower-ranked combination may be more practical.
> - **nSteps**: Higher values give finer search but exponentially increase runtime. Start with 10, increase only if needed.

**UDF Usage**
```
=UDF_REGRESS_OPTIMIZE(data, factorCols, resultCol, [goal], [topN], [nSteps])
```

#### FactorSweep

Single-factor What-If sweep. Fixes other factors at baseline values, sweeps one factor to analyze its impact on prediction results.

**VBA Usage**
```vb
FactorSweep(model, factorIndex, fromVal, toVal, steps, [baseValues]) As Variant()
```

| Parameter | Type | Description |
|------|------|------|
| model | Object | Model returned by LinearModelFit |
| factorIndex | Long | Factor index to sweep (1-based index in factorCols) |
| fromVal | Double | Start value |
| toVal | Double | End value |
| steps | Long | sweep steps |
| baseValues | Variant | Optional. Baseline value array for other factors |

```vb
Set model = LinearModelFit(dataRange, Array(1,2,3), 4)
Dim sweep As Variant
sweep = FactorSweep(model, 1, 100#, 140#, 11)
```

> **How to read the output**: Returns a 2D array with the swept factor's values and corresponding predicted results. Plots or scans the "what-if" curve — how does the predicted outcome change as one factor varies while others are held constant. Useful for visualizing marginal effects and finding optimal operating ranges for individual factors.

**UDF Usage**
```
=UDF_REGRESS_SWEEP(data, factorCols, resultCol, sweepFactorIdx, fromVal, toVal, steps)
```

---

> **Dependency chain**: RegressUtils -> LinearUtils + StatsUtils (TDist2T, FDistRT used for p value calculation; CorrelationMatrix used for UDF_REGRESS_CORREL). Import order: LinearUtils -> StatsUtils -> RegressUtils.

---

## Chapter 8: StringUtils — String Encoding, Validation, Distance & Formatting

String extraction, encoding conversion, format validation, edit distance, random generation, and text cleaning. **Module**: `StringUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`ExtractBetween`](#extractbetween) | `(text, startDelim, endDelim)` | Extract text between delimiters | String |
| [`ReverseString`](#reversestring) | `(text)` | Reverse string | String |
| [`CountSubstring`](#countsubstring) | `(text, substring, [caseSensitive])` | Count substring occurrences | Long |
| [`StartsWith`](#startswith) | `(text, prefix, [caseSensitive])` | Check prefixmatches | Boolean |
| [`EndsWith`](#endswith) | `(text, suffix, [caseSensitive])` | Check suffix matches | Boolean |
| [`LeftOf`](#leftof) | `(text, delimiter, [nth])` | Get text left of delimiter | String |
| [`RightOf`](#rightof) | `(text, delimiter, [nth], [fromRight])` | Get text right of delimiter | String |
| [`TextJoin`](#textjoin) | `(delimiter, arr)` | Join array with delimiter | String |
| [`NthWord`](#nthword) | `(text, n, [delimiter])` | Extract n-th word | String |
| [`CommonPrefix`](#commonprefix) | `(a, b, [caseSensitive])` | Longest common prefix of two strings | String |
| [`PadLeft`](#padleft) | `(text, totalWidth, [padChar])` | Pad left to specified width | String |
| [`PadRight`](#padleft) | `(text, totalWidth, [padChar])` | Pad right to specified width | String |
| [`Repeat`](#repeat) | `(text, count)` | Duplicate string n times | String |
| [`Truncate`](#truncate) | `(text, maxLen, [suffix])` | Truncate and append suffix | String |
| [`NormalizeWhitespace`](#normalizewhitespace) | `(text)` | Collapse extra whitespace characters | String |
| [`RemoveChars`](#removechars) | `(text, charsToRemove)` | Remove specified character set | String |
| [`KeepChars`](#removechars) | `(text, charsToKeep)` | Only preserve specified character set | String |
| [`ToTitleCase`](#totitlecase) | `(text)` | Smart title case (handles Mc/Mac/O') | String |
| [`RemoveDiacritics`](#removediacritics) | `(text)` | Remove diacritics (Unicode NFD) | String |
| [`Slugify`](#slugify) | `(text)` | generate URL-friendly slug | String |
| [`IsNullOrEmpty`](#isnullorempty) | `(value)` | check if Null/Empty/ZLS | Boolean |
| [`IsNullOrWhitespace`](#isnullorempty) | `(value)` | check if Null/Empty/whitespace-only | Boolean |
| [`IsEmail`](#isemail) | `(text)` | Email format validation | Boolean |
| [`IsUrl`](#isemail) | `(text)` | URL format validation | Boolean |
| [`Coalesce`](#coalesce) | `(ParamArray values())` | Return first non-empty value (like SQL COALESCE) | Variant |
| [`LevenshteinDistance`](#levenshteindistance) | `(s1, s2, [caseSensitive])` | Edit distance (Levenshtein) | Long |
| [`Soundex`](#soundex) | `(text)` | Phonetic hash (Soundex algorithm) | String |
| [`URLEncode`](#urlencode) | `(text)` | URL percent encode | String |
| [`URLDecode`](#urlencode) | `(text)` | URL percent decode | String |
| [`Base64Encode`](#base64encode) | `(text)` | Base64 encoding | String |
| [`Base64Decode`](#base64encode) | `(text)` | Base64 decoding | String |
| [`HTMLEncode`](#htmlencode) | `(text)` | HTML entity encoding | String |
| [`HTMLDecode`](#htmlencode) | `(text)` | HTML entity decoding | String |
| [`UUID`](#uuid) | `()` | Generate RFC 4122 v4 UUID | String |
| [`RandomString`](#randomstring) | `(length, [charset])` | Generate random string | String |

| Recipe | Functions Used |
|--------|---------------|
| [Text Cleaning](#recipe-text-cleaning) | `UDF_STR_NORMALIZEWHITESPACE`, `UDF_STR_REMOVECHARS`, `UDF_STR_TOTITLECASE` |
| [Encoding Conversion](#recipe-encoding) | `UDF_STR_BASE64ENCODE`, `UDF_STR_URLENCODE`, `UDF_STR_HTMLENCODE` |
| [Fuzzy Matching](#recipe-fuzzy-match) | `UDF_STR_LEVENSHTEIN`, `UDF_STR_SOUNDEX` |

<a id="recipe-text-cleaning"></a>
### Recipe 8.1 — Text Cleaning

**Scenario**: Data imported from external systems contains extra whitespace, special symbols, and inconsistent case; needs normalization before entering the library.

`=UDF_STR_NORMALIZEWHITESPACE(UDF_STR_TOTITLECASE(UDF_STR_REMOVECHARS(A1, "™®")))`

<a id="removechars"></a>
<a id="keepchars"></a>
#### RemoveChars / KeepChars

RemoveChars removes all characters from the string that are in the specified character set; KeepChars preserves only characters in the specified character set.

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

Collapses consecutive whitespace characters (spaces/tabs/newlines) into a single space, and trims leading/trailing.

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

Smart title case conversion: capitalizes first letter of each word, correctly handling Mc/Mac/O' prefixes.

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

Removes diacritics based on Unicode NFD decomposition (e.g. é→e, ü→u, ñ→n).

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

Converts text to URL-friendly format: lowercase + remove diacritics + replace non-alphanumeric with hyphens.

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
### Recipe 8.2 — Encoding Conversion

**Scenario**: Need to encode text as Base64 for transmission in JSON/XML, or encode URL parameters.

```
=UDF_STR_BASE64ENCODE(UDF_STR_URLENCODE(A1))
```

<a id="base64encode"></a>
<a id="base64decode"></a>
#### Base64Encode / Base64Decode

Standard Base64 encode/decode. Encoding input as ANSI/UTF-8 bytes; decoding returns original string.

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

Standard URL percent-encoding/decoding. Preserves alphanumeric and `-_.~`, encodes other characters as `%XX`.

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

HTML entity encode/decode: converts `<>&"` and other special characters to/from `&lt;` `&gt;` `&amp;` `&quot;`.

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
### Recipe 8.3 — Fuzzy Matching

**Scenario**: Need to match names between two customer rosters based on similarity, tolerating spelling differences.

```
=UDF_STR_LEVENSHTEIN(A1, B1)
=UDF_STR_SOUNDEX(A1)
```

#### LevenshteinDistance

Calculates edit distance between two strings (minimum number of insert/delete/replace operations). 0 means identical.

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

Generates a 4-character Soundex hash code based on English pronunciation, used for fuzzy matching phonetically similar words.

```vb
Soundex(text) As String
```

```vb
s = Soundex("Robert")  ' → "R163"
s = Soundex("Rupert")  ' → "R163" (Same phonetic group as Robert)
```

**UDF Usage**
```
=UDF_STR_SOUNDEX(text)
```

#### ExtractBetween

Extract text between two delimiters (excluding delimiters themselves).

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

Reverse string order.

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

Count the number of (non-overlapping) occurrences of a substring in text.

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

Check if string starts with a given prefix or ends with a given suffix.

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

Returns substring to the left or right of the first occurrence of delimiter. Returns empty string when delimiter not found.

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

Join array with delimiter, compatible with older Excel (no TEXTJOIN).

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

Extract n-th word by delimiter (1-based).

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

Returns longest common prefix of two strings. Default case-insensitive.

```vb
CommonPrefix(a, b, [caseSensitive]) As String
```

```vb
s = CommonPrefix("flower", "flow")     ' → "flow"
s = CommonPrefix("flower", "flight")   ' → "fl"
s = CommonPrefix("abc", "ABC")         ' → "abc"  (case-insensitive)
```

**UDF Usage**
```
=UDF_STR_COMMONPREFIX(a, b, [caseSensitive])
```

<a id="padleft"></a>
<a id="padright"></a>
#### PadLeft / PadRight

Pad string to target width with specified character (default space). Returns original text when target width is less than text length.

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

Truncate text to specified length, overflow replaced with optional suffix (default "...").

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

Empty/null value checks: IsNullOrEmpty checks Null/Empty/zero-length strings; IsNullOrWhitespace additionally rejects whitespace-only characters.

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

Simple format validation: IsEmail checks `@` and domain structure; IsUrl checks protocol prefix and domain format.

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

Returns the first non-empty value from the parameter list (not Null/Empty/ZLS), similar to SQL COALESCE.

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

UUID generates RFC 4122 v4 UUID; RandomString generates random string from specified character set; Repeat duplicates string n times.

```vb
UUID() As String
RandomString(length, [charset]) As String
Repeat(text, count) As String
```

```vb
s = UUID()                            ' → "550e8400-e29b-41d4-a716-446655440000"
s = RandomString(8)                   ' → "aB3xK9mQ" (defaultalphanumeric)
s = Repeat("ab", 3)                   ' → "ababab"
```

**UDF Usage**
```
=UDF_STR_UUID()
=UDF_STR_RANDOMSTRING(length, [charset])
=UDF_STR_REPEAT(text, count)
```

---

## Chapter 9: RegexUtils — Regex Match, Replace, Split & Capture Groups

Regex tools built on VBScript.RegExp. Supports match testing, extraction, replacement, splitting, counting, and capture groups. **Module**: `RegexUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`RegexIsMatch`](#regexismatch) | `(text, pattern, [ignoreCase], [multiLine])` | Test whether a match exists | Boolean |
| [`RegexExtract`](#regexextract) | `(text, pattern, [group], [ignoreCase])` | Extract first match (can specify capture group) | String |
| `RegexFindInRange` | `(rng, pattern, [ignoreCase])` | Find all matches within Range (VBA-only) | Range |
| [`RegexExtractAll`](#regexextractall) | `(text, pattern, [ignoreCase])` | Return all matches as array | String() |
| [`RegexExtractGroups`](#regexextractgroups) | `(text, pattern, [ignoreCase])` | Extract capture groups as 2D array | Variant(,) |
| [`RegexIsFullMatch`](#regexisfullmatch) | `(text, pattern, [ignoreCase])` | Test whether entire string matches | Boolean |
| [`RegexReplace`](#regexreplace) | `(text, pattern, replacement, [ignoreCase])` | Regex replace all matches | String |
| [`RegexReplaceRange`](#regexreplacerange) | `(rng, pattern, replacement, [ignoreCase])` | Row-wise regex replacement on range **Sub** | — |
| [`RegexSplit`](#regexsplit) | `(text, pattern, [ignoreCase])` | Regex split text | String() |
| [`RegexSplitToRange`](#regexreplacerange) | `(text, pattern, destCell, [ignoreCase])` | Split then output to worksheet **Sub** | — |
| [`RegexCount`](#regexcount) | `(text, pattern, [ignoreCase])` | Count matches | Long |
| [`RegexEscape`](#regexescape) | `(pattern)` | Escape regex special characters | String |

| Recipe | Functions Used |
|--------|---------------|
| [Regex Extraction](#recipe-regex-extract) | `UDF_REGEX_EXTRACT`, `UDF_REGEX_EXTRACTALL` |
| [Regex Replace](#recipe-regex-replace) | `UDF_REGEX_REPLACE` |
| [Text Splitting](#recipe-text-split) | `UDF_REGEX_SPLIT` |

<a id="recipe-regex-extract"></a>
### Recipe 9.1 — Regex Extraction

**Scenario**: Extract phone numbers, email addresses, or specific format IDs from free text.

```
=UDF_REGEX_EXTRACT(A1, "\d{3}-\d{4}-\d{4}")
=UDF_REGEX_EXTRACTALL(A1, "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
```

#### RegexIsMatch

Test whether the regex pattern matches anywhere in the text. Returns Boolean, does not extract content.

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

Extract the first match content. Use the `group` parameter to return a specific capture group (0 = full match).

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

Return all non-overlapping matches as a string array.

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

Extract all capture groups from all matches, returning a 2D array: one match per row, first capture group per column.

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

Test whether the entire string fully matches the regex pattern (equivalent to `^pattern$`).

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
### Recipe 9.2 — Regex Replace

**Scenario**: Batch-clean data: remove HTML tags, standardize date formats, mask phone numbers.

```
=UDF_REGEX_REPLACE(A1, "<[^>]+>", "")
=UDF_REGEX_REPLACE(A1, "(\d{3})\d{4}(\d{4})", "$1****$2")
```

#### RegexReplace

Replace all matches of the regex pattern in text. Supports `$1`-`$9` backreferences.

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

Perform regex replacement on each cell in a worksheet range, modifying in-place.

```vb
RegexReplaceRange(rng, pattern, replacement, [ignoreCase])
```

```vb
RegexReplaceRange Range("A1:A100"), "^\s+|\s+$", ""
```

<a id="recipe-text-split"></a>
### Recipe 9.3 — Text Splitting

**Scenario**: Split a text column by multiple delimiters, such as comma/semicolon/pipe mixed-separator data.

```
=UDF_REGEX_SPLIT(A1, "[,;|]\s*")
```

#### RegexSplit

Split text by a regex pattern, returning a string array. More flexible than VBA `Split()` (supports multi-character delimiters, empty/null tolerance).

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

Count non-overlapping matches of the regex pattern in text.

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

Escape regex special characters in a string so they are matched literally.

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

## Chapter 10: JsonUtils — Pure VBA JSON Parsing, Query & Serialization

Pure VBA recursive-descent JSON parser with zero external dependencies. Supports objects, arrays, strings, numbers, booleans, null, nesting, and full Unicode (including surrogate pairs). **Module**: `JsonUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`JsonParse`](#jsonparse) | `(jsonText)` | Parse JSON string **VBA-only** | Variant (Dictionary/Array/Scalar) |
| [`JsonGet`](#jsonget) | `(json, path)` | Extract by pathvalue **VBA-only** | Variant |
| [`JsonToRange`](#jsontorange) | `(json, destCell, [headers])` | JSON object array output to worksheet **Sub** | — |
| [`JsonGetKeys`](#jsongetkeys) | `(json)` | List all keys of an object | String() |
| [`JsonIsValid`](#jsonisvalid) | `(jsonText)` | Validate JSON validity (no errors thrown) | Boolean |
| [`JsonStringify`](#jsonstringify) | `(value)` | VBA Variant → JSON string **VBA-only** | String |

| Recipe | Functions Used |
|--------|---------------|
| [JSON path extraction](#recipe-json-path) | `UDF_JSON_GET` |
| [Table to JSON](#recipe-table-to-json) | `UDF_JSON_STRINGIFY` |

<a id="recipe-json-path"></a>
### Recipe 10.1 — JSON Path Extraction

**Scenario**: Extract nested field values by path from API-returned JSON responses.

```
=UDF_JSON_GET(A1, "users[0].name")
=UDF_JSON_GET(A1, "data.items[2].price")
```

#### JsonParse

Parse a JSON string into VBA native types. Object → Dictionary, arrays → Variant(), scalars converted by type. **VBA-only** (Dictionary cannot be returned directly to a worksheet). Object keys are case-sensitive (RFC 8259): `{"a":1,"A":2}` keeps both keys.

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

Extract values from parsed JSON by dot/bracket path. Supports `a.b`, `arr[0]`, `users[0].name`. **VBA-only**.

**VBA Usage**
```vb
JsonGet(json, path) As Variant
```

```vb
Dim data As Variant: data = JsonParse("{""users"":[{""name"":""Alice""}]}")
v = JsonGet(data, "users[0].name")  ' → "Alice"
```

#### UDF_JSON_GET

Worksheet version of JsonGet. Objects/arrays return placeholder text `[Object]` / `[Array]`, scalars return the actual value.

**UDF Usage**
```
=UDF_JSON_GET(jsonText, path)
```

| Parameter | Type | Description |
|------|------|------|
| jsonText | String | JSON string |
| path | String | Dot-path, e.g. `data.items[0].name` |

#### JsonGetKeys

Return all top-level keys of a JSON object as a string array.

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

Validate that a string is valid JSON, without throwing exceptions.

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
### Recipe 10.2 — Table to JSON

**Scenario**: Serialize a worksheet range into a JSON string, for API request bodies or config files.

```
=UDF_JSON_STRINGIFY(range)
```

#### JsonStringify

Serialize a VBA Variant (Dictionary/Array/Scalar) into a JSON string. **VBA-only**.

**VBA Usage**
```vb
JsonStringify(value) As String
```

```vb
s = JsonStringify(Array(1, 2, 3))  ' → "[1,2,3]"
```

#### JsonToRange (Sub)

Output a JSON object array to a worksheet, auto-generating table headers.

**VBA Usage**
```vb
JsonToRange(jsonText, destCell, [includeHeaders])
```

```vb
JsonToRange "[{""a"":1,""b"":2},{""a"":3,""b"":4}]", Range("A1")
```

---

## Chapter 11: XmlUtils — XML Parsing & Tabulation

XML tools based on MSXML2 (Windows built-in). XPath queries, attribute extraction, XML to worksheet. **Module**: `XmlUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|----------|-----------|-------------|---------|
| [`XmlValidate`](#xmlvalidate) | `(xml, [errDetail])` | Validate XML format | Boolean |
| [`XmlGet`](#xmlget) | `(xml, xpath)` | XPath query value | Variant |
| [`XmlGetAttr`](#xmlgetattr) | `(xml, xpath, attrName)` | Get attribute value | Variant |
| [`XmlToRange`](#xmltorange) | `(xml, rowXPath, [colNames])` | XML → 2D arrays | Variant |

| Recipe | Functions Used |
|--------|---------------|
| [XPath data extraction](#recipe-xpath-extract) | `UDF_XML_GET`, `XmlGetAttr` |
| [XML to table](#recipe-xml-to-table) | `UDF_XML_TABLE`, `XmlToRange` |

### Function Details

#### XmlValidate

Validate XML format. **VBA-only** version provides error details (`errDetail` output parameter).

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

XPath query for node text. Supports `/a/b`, `//c`, `[@k='v']`, `[1]`.

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

`=UDF_XML_GET(A1, "/root/name")` → node text

#### XmlGetAttr

Get the attribute value of an XPath node. **VBA-only**.

```vb
XmlGetAttr(xml, xpath, attrName) As Variant
```

```vb
XmlGetAttr("<user id='007'/>", "/user", "id")  ' → "007"
```

#### XmlToRange

XML repeating nodes → 2D array (can be written directly to a Range). Auto-detect child element column names.

```vb
XmlToRange(xml, rowXPath, [colNames]) As Variant
```

| Parameter | Type | Description |
|-----------|------|-------------|
| xml | String | XML text |
| rowXPath | String | Row node XPath |
| colNames | Variant | (Optional) Column name array; auto-detect if omitted |

```vb
Dim arr As Variant
arr = XmlToRange(xml, "/rows/row", Array("colA", "colB"))
ws.Range("A1").Resize(UBound(arr,1), UBound(arr,2)).Value = arr
```

**UDF Usage**
```
=UDF_XML_TABLE(xml, rowXPath, [colNames])
```

`=UDF_XML_TABLE(A1, "/rows/row", {"a","b"})` → 2D spill array

---


<a id="recipe-xpath-extract"></a>

### Recipe 11.1 — XPath data extraction

**Scenario**: Extract specified node text or attribute values from XML using XPath.

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

### Recipe 11.2 — XML to table

**Scenario**: Import a list of XML repeating elements directly into an Excel worksheet.

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

## Chapter 12: DateTimeUtils — ISO Weeks, Workdays, Age & Date Ranges

Date information extraction, workday calculations, Unix timestamp conversion, age, and holidays. **Module**: `DateTimeUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`ISOWeekNum`](#isoweeknum) | `(dt)` | ISO 8601 week number (1-53) | Long |
| [`FirstDayOfMonth`](#firstdayofmonth) | `(dt)` | First day of month | Date |
| [`LastDayOfMonth`](#lastdayofmonth) | `(dt)` | Last day of month | Date |
| [`FirstDayOfQuarter`](#firstdayofquarter) | `(dt)` | First day of quarter | Date |
| [`LastDayOfQuarter`](#lastdayofquarter) | `(dt)` | Last day of quarter | Date |
| [`DaysInMonth`](#daysinmonth) | `(y, m)` | Days in month | Long |
| [`DaysInYear`](#daysinyear) | `(y)` | Days in year (365/366) | Long |
| [`DayOfYear`](#dayofyear) | `(dt)` | Day of year (1-366) | Long |
| [`IsLeapYear`](#isleapyear) | `(year)` | Leap year check | Boolean |
| [`Quarter`](#quarter) | `(dt)` | Quarter (1-4) | Long |
| [`FiscalYear`](#fiscalyear) | `(dt, [startMonth])` | Fiscal year | Long |
| [`UnixToDate`](#unixtodate) | `(timestamp)` | Unix timestamp → Date | Date |
| [`DateToUnix`](#datetounix) | `(dt)` | Date → Unix timestamp | Double |
| [`IsWeekend`](#isweekend) | `(dt)` | Check if Saturday/Sunday | Boolean |
| [`IsHoliday`](#isholiday) | `(dt, [holidays])` | Check if holiday | Boolean |
| [`AddMonthsSafe`](#addmonthssafe) | `(dt, months)` | Safe add months (auto-clamp at month end) | Date |
| [`WorkdaysBetween`](#workdaysbetween) | `(startDate, endDate, [holidays])` | Workdays between two dates | Long |
| [`NextWorkday`](#nextworkday) | `(startDate, days, [holidays])` | Date after n workdays | Date |
| [`WeekOfMonth`](#weekofmonth) | `(dt)` | Week of month (1-6) | Long |
| [`StartOfWeek`](#startofweek) | `(dt, [firstDayOfWeek])` | First day of week containing date | Date |
| [`EndOfWeek`](#endofweek) | `(dt, [firstDayOfWeek])` | Last day of week containing date | Date |
| [`DateRange`](#daterange) | `(startDate, endDate)` | Generate sequential date array | Date() |
| [`Age`](#age) | `(birthDate, [asOf])` | Detailed age (years/months/days) **VBA-only** | Dictionary |
| [`AgeYears`](#ageyears) | `(birthDate, [refDate])` | Simple integer age | Long |
| [`NthWeekday`](#nthweekday) | `(year, month, weekday, n)` | The nth specified weekday of a month | Date |
| [`Easter`](#easter) | `(year)` | Easter date (Anonymous algorithm) | Date |
| [`DateDiffParts`](#datediffparts) | `(d1, d2)` | Precise date difference **VBA-only** | Dictionary |

| Recipe | Functions Used |
|--------|---------------|
| [Date information extraction](#recipe-date-info) | `UDF_DT_ISOWEEKNUM`, `UDF_DT_QUARTER`, `UDF_DT_ISLEAPYEAR` |
| [Workday Calculation](#recipe-workday) | `UDF_DT_WORKDAYSBETWEEN`, `UDF_DT_NEXTWORKDAY` |
| [Unix Timestamp](#recipe-unix-timestamp) | `UDF_DT_UNIXTODATE`, `UDF_DT_DATETOUNIX` |

<a id="recipe-date-info"></a>
### Recipe 12.1 — Date information extraction

**Scenario**: Extract ISO week number, quarter, and leap year flags from a date column for group summarization.

```
=UDF_DT_ISOWEEKNUM(A1)
=UDF_DT_QUARTER(A1)
=UDF_DT_ISLEAPYEAR(YEAR(A1))
```

#### ISOWeekNum

Return the ISO 8601 week number (1-53). ISO week definition: Monday is the first day of the week, and the week containing the first Thursday of the year is week 1.

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
<a id="firstdayof/lastdayof month"></a>
<a id="quarter"></a>
#### FirstDayOf/LastDayOf Month / Quarter

Return the first day or last day of the month or quarter containing the given date.

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

Date component extraction functions.

```vb
DaysInMonth(dt) As Long       ' Days in month
DaysInYear(dt) As Long        ' Days in year
DayOfYear(dt) As Long         ' day of year (1-366)
IsLeapYear(year) As Boolean   ' Leap year check
Quarter(dt) As Long           ' Quarter 1-4
FiscalYear(dt, [startMonth]) As Long  ' Fiscal year (default January start)
```

```vb
n = DaysInMonth(2024, 2)                 ' → 29
n = Quarter(#2024-07-15#)               ' → 3
n = FiscalYear(#2025-12-15#, 7)         ' → 2026 (July start)
```

**UDF Usage**
```
=UDF_DT_DAYSINMONTH(year, month)   =UDF_DT_DAYSINYEAR(year)
=UDF_DT_DAYOFYEAR(date)     =UDF_DT_ISLEAPYEAR(year)
=UDF_DT_QUARTER(date)       =UDF_DT_FISCALYEAR(date, [startMonth])
```

#### AgeYears

Calculate integer age from birth date to reference date.

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

Calculate detailed age (years, months, days) from birth date to reference date. Returns a Dictionary with "years", "months", "days" keys. **VBA-only**.

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

WeekOfMonth returns the week number within the month (1-6); StartOfWeek/EndOfWeek return the start/end dates of the week containing the given date.

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

NthWeekday returns the date of the nth specified weekday in a given year and month; Easter returns the Easter date.

```vb
NthWeekday(year, month, weekday, n) As Date
Easter(year) As Date
```

```vb
d = NthWeekday(2025, 6, vbMonday, 2)  ' → 2025-06-09 (2nd Monday of June)
d = Easter(2025)                        ' → 2025-04-20
```

**UDF Usage**
```
=UDF_DT_NTHWEEKDAY(year, month, weekday, n)
=UDF_DT_EASTER(year)
```

#### DateDiffParts

Calculate the precise difference between two dates, returning a Dictionary containing years/months/days/totalDays. **VBA-only**.

```vb
DateDiffParts(d1, d2) As Object  ' → Dictionary
```

```vb
Set diff = DateDiffParts(#2020-03-15#, #2025-06-11#)
' diff("years")=5, diff("months")=2, diff("days")=27
```

<a id="recipe-workday"></a>
### Recipe 12.2 — Workday Calculation

**Scenario**: Calculate project duration (workdays between) or determine the delivery date after n workdays.

```
=UDF_DT_WORKDAYSBETWEEN(A1, B1, holidaysRange)
=UDF_DT_NEXTWORKDAY(A1, 10, holidaysRange)
```

<a id="workdaysbetween"></a>
<a id="nextworkday"></a>
#### WorkdaysBetween / NextWorkday

WorkdaysBetween computes the number of workdays between two dates (Monday-Friday, inclusive); NextWorkday returns the date after n workdays. Supports an optional holidays list.

```vb
WorkdaysBetween(startDate, endDate, [holidays]) As Long
NextWorkday(startDate, days, [holidays]) As Date
```

```vb
n = WorkdaysBetween(#2025-06-09#, #2025-06-13#)  ' → 5
d = NextWorkday(#2025-06-11#, 3)                  ' → 2025-06-16 (skips weekends)
```

**UDF Usage**
```
=UDF_DT_WORKDAYSBETWEEN(startDate, endDate, [holidays])
=UDF_DT_NEXTWORKDAY(startDate, days, [holidays])
```

<a id="isweekend"></a>
<a id="isholiday"></a>
#### IsWeekend / IsHoliday

IsWeekend checks if the date is Saturday or Sunday; IsHoliday checks if it is in the specified holidays list.

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

Safely add months: if the target date exceeds the maximum days of the resulting month (e.g., Jan 31 + 1 month), it auto-clamps to the last day of the month.

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

Generate an array of all dates from the start date to the end date.

```vb
DateRange(startDate, endDate) As Date()
```

**UDF Usage**
```
=UDF_DT_DATERANGE(startDate, endDate)
```

<a id="recipe-unix-timestamp"></a>
### Recipe 12.3 — Unix Timestamp

**Scenario**: Exchange data containing Unix timestamps with external systems (API/data warehouse).

```
=UDF_DT_UNIXTODATE(A1)
=UDF_DT_DATETOUNIX(A1)
```

<a id="unixtodate"></a>
<a id="datetounix"></a>
#### UnixToDate / DateToUnix

Convert between Unix timestamps (seconds since UTC 1970-01-01) and VBA Date type. Conversion is UTC-based, without timezone offset.

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

## Chapter 13: RangeUtils — Range Export, Operations & Named Range Management

Worksheet range positioning, export (HTML/JSON/Markdown/CSV), filtering, sorting, deduplication & named range management. **Module**: `RangeUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`LastRow`](#lastrow) | `([ws])` | Last data row number in worksheet | Long |
| [`LastCol`](#lastrow) | `([ws])` | Last data column number in worksheet | Long |
| [`FirstRow`](#lastrow) | `([ws])` | First data row number | Long |
| [`FirstCol`](#lastrow) | `([ws])` | First data column number | Long |
| [`ColLetter`](#colletter) | `(colNum)` | Column number → letter | String |
| [`ColNumber`](#colletter) | `(colLetter)` | Letter → column number | Long |
| [`UsedRangeEx`](#lastrow) | `([ws])` | Actual used data range | Range |
| [`GetCellAddress`](#getcelladdress) | `(row, col, [absolute])` | Build cell address | String |
| [`RangeToHTML`](#rangetohtml) | `(rng, [includeHeaders])` | Range → HTML table | String |
| [`RangeToJSON`](#rangetohtml) | `(rng, [includeHeaders])` | range → JSON arrays | String |
| [`RangeToMarkdown`](#rangetohtml) | `(rng, [includeHeaders])` | Range → Markdown table | String |
| [`ExportRangeToCSV`](#exportrangetocsv) | `(rng, filePath, [delimiter])` | Export to CSV file **Sub** | -- |
| [`FilterRangeToArray`](#filterrangetoarray) | `(rng, filterCol, operator, value)` | Conditional filter, returns array | Variant(,) |
| [`FindAll`](#findall) | `(rng, what, [lookIn])` | Find all matching cells | Range |
| [`MergeRanges`](#mergeranges) | `(rng1, rng2)` | Get range union | Range |
| [`IntersectRanges`](#mergeranges) | `(rng1, rng2)` | Get range intersection | Range |
| [`RangeExists`](#rangeexists) | `(rangeName)` | Check if reference is #REF! | Boolean |
| [`NamedRangeExists`](#namedrange-operations) | `(name)` | Check if named range exists | Boolean |
| [`NamedRangeAdd`](#namedrange-operations) | `(name, rng, [scope])` | Create/update named range **Sub** | -- |
| [`NamedRangeDelete`](#namedrange-operations) | `(name, [scope])` | Delete named range **Sub** | -- |
| [`IsRangeEmpty`](#israngeempty) | `(rng)` | Whether range is all empty/null | Boolean |
| [`CountVisible`](#israngeempty) | `(rng)` | Visible row count | Long |
| [`RangeDiff`](#rangediff) | `(rng1, rng2)` | Compare range differences | Variant(,) |
| [`SafeText`](#safetext) | `(cell)` | Safely convert cell value to string | String |
| [`ValuesEqual`](#rangeexists) | `(v1, v2)` | Value equality comparison (tolerant) | Boolean |
| [`CopyRangeToSheet`](#lastrow) | `(arr, destCell)` | Array → worksheet **Sub** | -- |
| [`ClearRangeContents`](#clearrangecontents) | `(rng)` | Clear contents **Sub** | -- |
| [`ClearRangeFormats`](#clearrangecontents) | `(rng)` | Clear formats **Sub** | -- |
| [`RemoveDuplicatesRange`](#removeduplicatesrange) | `(rng, cols)` | Deduplicate by column **Sub** | -- |
| [`SortRange`](#sortrange) | `(rng, keyCol, [ascending])` | Sort by column **Sub** | -- |

| Recipe | Functions Used |
|--------|---------------|
| [Range Export to HTML](#recipe-range-export-html) | `UDF_RANGE_TOHTML` |
| [Range Export to JSON](#recipe-range-export-json) | `UDF_RANGE_TOJSON` |
| [Conditional Filter](#recipe-conditional-filter) | `UDF_RANGE_FILTER` |

<a id="recipe-range-export-html"></a>
### Recipe 13.1 -- Range Export to HTML

**Scenario**: Quickly export a data table to an HTML fragment for use in email bodies or web embedding.

```
=UDF_RANGE_TOHTML(A1:D10, TRUE)
```

<a id="rangetohtml"></a>
<a id="rangetojson"></a>
<a id="rangetomarkdown"></a>
<a id="rangetohtml"></a>
<a id="rangetojson"></a>
<a id="rangetomarkdown"></a>
#### RangeToHTML / RangeToJSON / RangeToMarkdown

Export a range to an HTML table, JSON array, or Markdown table string. Optionally includes the header row.

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

Export a range to a CSV file. Supports custom delimiter.

```vb
ExportRangeToCSV rng, filePath, [delimiter]
```

```vb
ExportRangeToCSV Range("A1:D100"), "C:\Data\export.csv", ";"
```

#### SafeText

Safely get cell text value. Error values (#N/A, #VALUE! etc.) return their display text.

```vb
SafeText(cell) As String
```

**UDF Usage**
```
=UDF_RANGE_SAFETEXT(cell)
```

#### GetCellAddress

Build a cell address string for a specified row and column, supports absolute/relative reference.

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
### Recipe 13.2 — Range Export to JSON

**Scenario**: Serialize table data to JSON for Web API data exchange.

```
=UDF_RANGE_TOJSON(A1:D10, TRUE)
```

<a id="lastrow"></a>
<a id="lastcol"></a>
<a id="firstrow"></a>
<a id="firstcol"></a>
<a id="lastrow"></a>
<a id="lastcol"></a>
<a id="firstrow"></a>
<a id="firstcol"></a>
#### LastRow / LastCol / FirstRow / FirstCol

Returns the data boundary row/column numbers of a worksheet. LastRow/LastCol scan upward/leftward from the last row/column.

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

Convert between column number and column letter: ColLetter(1) → "A", ColNumber("Z") → 26.

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
<a id="rangeexists"></a>
<a id="israngeempty"></a>
<a id="countvisible"></a>
#### RangeExists / IsRangeEmpty / CountVisible

RangeExists checks reference validity (not #REF!); IsRangeEmpty checks whether a range is entirely empty/null; CountVisible counts visible rows.

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

Named range management: NamedRangeAdd creates or updates a named range; NamedRangeDelete deletes a named range; NamedRangeExists checks if a named range exists.

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
### Recipe 13.3 -- Conditional Filter

**Scenario**: Filter range data by a column condition; the result is returned as an array for further processing.

```
=UDF_RANGE_FILTER(A1:D100, 2, ">=", 80)
```

#### FilterRangeToArray

Filter a range by a column, returning a 2D array of rows that meet the condition. Supports `=`, `<`, `>`, `<=`, `>=`, `<>`, `contains`, `regex`.

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

> **Known Limitation**: In legacy Excel this UDF does not produce a full spill array (requires Ctrl+Shift+Enter array formula or Excel 365 dynamic arrays); prefer calling `FilterRangeToArray` directly from VBA.

#### FindAll

Find all matching cells in a range, returns a Union Range. Supports searching by value/formula/comment.

```vb
FindAll(rng, what, [lookIn]) As Range
```

```vb
Set found = FindAll(Range("A1:A100"), "Yes")
```

<a id="mergeranges"></a>
<a id="intersectranges"></a>
#### MergeRanges / IntersectRanges

MergeRanges returns the union of two ranges; IntersectRanges returns their intersection.

```vb
MergeRanges(rng1, rng2) As Range
IntersectRanges(rng1, rng2) As Range
```

#### RangeDiff

Compare value differences between two ranges, returns a table with row/column and content comparison of differing cells.

```vb
RangeDiff(rng1, rng2, [outTruncated]) As Range
```

<a id="clearrangecontents"></a>
<a id="clearrangeformats"></a>
<a id="removeduplicatesrange"></a>
<a id="sortrange"></a>
#### Sub Procedures (Clear/Sort/RemoveDuplicates/Copy)

In-place operation subroutines:

```vb
ClearRangeContents rng           ' clear contents
ClearRangeFormats rng            ' Clear formats
AutoFitRange rng                 ' Auto-fit column width/row height
SortRange rng, keyCol, [ascending]   ' Sort by column
RemoveDuplicatesRange rng, cols  ' Deduplicate by specified columns
CopyRangeToSheet arr, destCell   ' Write array to worksheet
```

```vb
RemoveDuplicatesRange Range("A1:D100"), Array(1, 2)  ' Deduplicate by columns 1, 2
CopyRangeToSheet myArray, Range("A1")
```

---

## Chapter 14: FileSystemUtils — UTF-8 File I/O, Folder Traversal & Drive Info

File system tools based on Scripting.FileSystemObject and ADODB.Stream, supporting UTF-8 text I/O, file/folder enumeration, path parsing & drive information. **Module**: `FileSystemUtils.bas`

> **⚠️ Path Safety Restrictions**: Every path argument passes `ValidateSafePath`: paths containing `..` (directory traversal) and UNC paths (`\\server\share` as well as `//server/share`) are rejected. Known limitation: symbolic links/junctions are not resolved. Note the failure semantics: except for `ReadBinaryFile` (which raises), most functions fail silently when a path is rejected (e.g. `ReadTextFile` returns `""`, `DeleteFile`/`CopyFileSafe` return `False`) — callers previously using UNC or `..` paths will see "silent success/empty results" after upgrading; audit your paths first.

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`NormalizePath`](#normalizepath) | `(path)` | Normalize delimiter to `\` | String |
| [`PathCombine`](#normalizepath) | `(path1, path2)` | Safely join paths | String |
| [`IsPathValid`](#ispathvalid) | `(path)` | Path syntax is valid | Boolean |
| [`FileExists`](#fileexists) | `(filePath)` | Check if file exists | Boolean |
| [`FolderExists`](#fileexists) | `(folderPath)` | Check if folder exists | Boolean |
| [`GetFileName`](#getfilename) | `(path)` | Extract filename (with extension) | String |
| [`GetBaseName`](#getfilename) | `(path)` | Extract filename (without extension) | String |
| [`GetExtension`](#getfilename) | `(path)` | Extract extension | String |
| [`GetFolderPath`](#getfolderpath) | `(path)` | Extract parent folder path | String |
| [`GetFileSize`](#getfilesize) | `(filePath)` | File size (bytes) | Double |
| [`GetFileSizeFmt`](#getfilesize) | `(filePath)` | Format file size | String |
| [`FileModified`](#filemodified) | `(filePath)` | File last modified time | Date |
| [`ReadTextFile`](#readtextfile) | `(filePath, [encoding])` | Read text content | String |
| [`WriteTextFile`](#writetextfile) | `(filePath, text, [encoding])` | Write text file **Sub** | -- |
| [`AppendTextFile`](#appendtextfile) | `(filePath, text, [encoding])` | Append text **Sub** | -- |
| [`ListFiles`](#listfiles) | `(folderPath, [pattern])` | List file path array | String() |
| [`ListFolders`](#listfiles) | `(folderPath)` | List subfolder path array | String() |
| [`EnsureFolder`](#ensurefolder) | `(folderPath)` | Recursively create directory | Boolean |
| [`TempFileName`](#tempfilename) | `([prefix], [extension])` | Generate temp filename | String |
| [`GetTempFolder`](#tempfilename) | `()` | System temp folder path | String |
| [`GetSpecialFolder`](#getspecialfolder) | `(folderType)` | Special folder path | String |
| [`DeleteFile`](#deletefile) | `(filePath)` | Delete file | Boolean |
| [`CopyFileSafe`](#deletefile) | `(sourcePath, destPath)` | Safe copy (auto-create directory) | Boolean |
| [`ReadBinaryFile`](#readbinaryfile) | `(filePath)` | Read binary file | Byte() |
| [`WriteBinaryFile`](#readbinaryfile) | `(filePath, data)` | Write binary file **Sub** | -- |
| [`GetDriveInfo`](#getdriveinfo) | `(driveLetter)` | Drive information **VBA-only** | Dictionary |
| [`CopyFolder`](#copyfolder) | `(source, dest)` | Recursively copy folder **Sub** | -- |
| [`DeleteFolder`](#copyfolder) | `(source)` | Recursively delete folder **Sub** | -- |

| Recipe | Functions Used |
|--------|---------------|
| [Batch Text File Reading](#recipe-batch-file-read) | `UDF_FS_READTEXT`, `UDF_FS_LISTFILES` |
| [Path Parsing](#recipe-path-parse) | `UDF_FS_FILENAME`, `UDF_FS_BASENAME`, `UDF_FS_EXTENSION` |
| [File Information](#recipe-file-info) | `UDF_FS_FILESIZE`, `UDF_FS_FILEMODIFIED`, `UDF_FS_FILEEXISTS` |

<a id="recipe-batch-file-read"></a>
### Recipe 14.1 — Batch Text File Reading

**Scenario**: Read all log files in a folder for combined analysis.

> **Note**: `UDF_FS_LISTFILES` returns an array of paths. To read multiple files at once, enter the formula as a dynamic array formula (Excel 365) or use Ctrl+Shift+Enter (legacy Excel). For single-file reading, use `UDF_FS_READTEXT` with a direct path.

```
=UDF_FS_READTEXT(UDF_FS_LISTFILES("C:\Logs", "*.log"))
```

#### ReadTextFile

Read text file content. Default UTF-8 encoding. Supported encodings: `utf-8`, `ansi`, `unicode`.

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
<a id="appendtextfile (sub)"></a>
#### WriteTextFile / AppendTextFile (Sub)

Write or append text content. Default UTF-8 encoding with BOM.

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

Read or write raw byte content of a file.

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

ListFiles lists file paths matching a pattern within a folder; ListFolders lists subfolder paths.

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

Recursively create folder path, skip if it already exists. Returns success status.

```vb
EnsureFolder(folderPath) As Boolean
```

**UDF Usage**
```
=UDF_FS_ENSUREFOLDER(folderPath)
```

<a id="recipe-path-parse"></a>
### Recipe 14.2 — Path Parsing

**Scenario**: Extract filename, base name, or extension from a full file path.

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

Extract components from a full path.

```vb
GetFileName(path) As String     ' → "report.xlsx"
GetBaseName(path) As String     ' → "report"
GetExtension(path) As String    ' → ".xlsx" (includes dot)
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

NormalizePath unifies path delimiters to backslash; PathCombine safely joins two path segments (handles duplicate delimiters).

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

Check if a path string syntax is valid (no invalid characters, correct format).

```vb
IsPathValid(path) As Boolean
```

**UDF Usage**
```
=UDF_FS_ISPATHVALID(path)
```

<a id="recipe-file-info"></a>
### Recipe 14.3 -- File Information

**Scenario**: List files with size, modification time, and existence information.

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

File information query functions.

```vb
GetFileSize(filePath) As Double          ' byte count
GetFileSizeFmt(filePath) As String       ' formatted, e.g. "1.5 MB"
FileModified(filePath) As Date           ' last modified time
FileExists(filePath) As Boolean          ' File exists
FolderExists(folderPath) As Boolean      ' folderexists
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

DeleteFile deletes a file (returns True on success); CopyFileSafe copies a file and auto-creates the target directory.

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

TempFileName generates a unique filename in the temp directory; GetTempFolder returns the system temp directory; GetSpecialFolder returns Desktop/Documents and other special folders.

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

Returns drive information Dictionary containing TotalSize/FreeSpace/FileSystem/VolumeName/IsReady. **VBA-only**.

```vb
GetDriveInfo(driveLetter) As Object  ' → Dictionary
```

```vb
Set info = GetDriveInfo("C")
' info("TotalSize") → 256060514304
' info("FreeSpace") → 128030257152
```

<a id="copyfolder"></a>
<a id="deletefolder (sub)"></a>
#### CopyFolder / DeleteFolder (Sub)

Recursively copy or delete an entire folder.

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

## Chapter 15: PhyChemUtils — Molecular Weight, Unit Conversion & Gas Equations of State

Physicochemical calculation tools: chemical formula molecular weight, unit conversion (volume/pressure/temperature/mass), solution dilution, ideal gas & Peng-Robinson real gas. **Module**: `PhyChemUtils.bas`

**Quick Reference**

| Function | Parameters | Description | Returns |
|------|------|------|--------|
| [`MolecularWeight`](#molecularweight) | `(formula)` | Chemical formula → Molar mass (g/mol) | Double |
| [`ConvertMass`](#convertmass) | `(val, fromUnit, [toUnit])` | Mass unit conversion | Double |
| [`ConvertVolume`](#convertvolume) | `(val, fromUnit, [toUnit])` | Volume unit conversion | Double |
| [`ConvertPressure`](#convertpressure) | `(val, fromUnit, [toUnit])` | Pressure unit conversion | Double |
| [`ConvertTemperature`](#converttemperature) | `(val, fromUnit, [toUnit])` | Temperature unit conversion | Double |
| [`ConvertStandard`](#convertstandard) | `(v, p, T, MW)` | Gas standard state volume and mass | Variant(,) |
| [`MassToMoles`](#masstomoles) | `(mass, molWeight)` | Mass (g) → amount of substance (mol) | Double |
| [`MolesToMass`](#masstomoles) | `(moles, molWeight)` | Amount of substance (mol) → mass (g) | Double |
| [`DilutionSolve`](#dilutionsolve) | `(c1, v1, c2, v2)` | C1V1=C2V2 solver (solve for missing) | Double |
| [`IdealGasLaw`](#idealgaslaw) | `(P, V, n, T)` | PV=nRT solver (solve for missing) | Double |
| [`Density`](#density) | `(m, v, rho)` | Density solver (solve for missing) | Double |
| [`PercentYield`](#percentyield) | `(actual, theoretical)` | Yield percentage | Double |
| [`CompressFactorPR`](#compressfactorpr) | `(P, T, Tc, Pc, omega)` | Peng-Robinson compressibility factor Z | Double |
| [`CylinderStdVolume`](#cylinderstdvolume) | `(cylVol, fillP, fillT, gasName)` | Cylinder standard state volume (with Z factor) | Double |
| [`CylinderStdVolumeFromMass`](#cylinderstdvolumefrommass) | `(netWt, gasFormula)` | Cylinder standard state volume (known net weight) | Double |

| Recipe | Functions Used |
|--------|---------------|
| [Molecular Weight & Dilution](#recipe-molweight-dilution) | `UDF_PC_MOLWEIGHT`, `UDF_PC_DILUTION` |
| [Gas Conversion](#recipe-gas-conversion) | `UDF_PC_CONVERTSTANDARD`, `UDF_PC_STDVOLUME` |
| [Unit Conversion](#recipe-unit-conversion) | `UDF_PC_CONVERTMASS`, `UDF_PC_CONVERTVOLUME`, `UDF_PC_CONVERTPRESSURE`, `UDF_PC_CONVERTTEMPERATURE` |

<a id="recipe-molweight-dilution"></a>
### Recipe 15.1 -- Molecular Weight & Dilution

**Scenario**: Auto-calculate molar mass from a chemical formula, or solve dilution ratios using C1V1=C2V2.

```
=UDF_PC_MOLWEIGHT("H2SO4")
=UDF_PC_DILUTION(10, 100, , 500)     ' Solve for v2 → 200 mL
```

#### MolecularWeight

Auto-calculate molar mass (g/mol) from a chemical formula. Supports parentheses, nesting, and 103 common elements from the periodic table.

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

C1 * V1 = C2 * V2 solver. Pass three of the four parameters; pass 0 or omit the parameter to solve for. Returns the solved value.

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

Convert between mass and amount of substance.

```vb
MassToMoles(mass, molWeight) As Double
MolesToMass(moles, molWeight) As Double
```

Unit convention: mass in g, moles in mol, molWeight in g/mol.

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

Density solver: calculate the third quantity from any two known quantities among m/V/rho.

```vb
Density(m, v, rho) As Double
```

Unit convention: mass in g, volume in mL, rho in g/mL.

```vb
v = Density(100, 0, 2.5)  ' → 40 (volume = mass/density)
```

**UDF Usage**
```
=UDF_PC_DENSITY(m, v, rho)
```

#### PercentYield

Calculate reaction yield percentage: actual yield / theoretical yield * 100.

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

Unit conversion functions. Each function accepts `(value, fromUnit, [toUnit])` parameters; omitting the target unit defaults to converting to the SI unit.

Supported unit options (the first parameter value auto-detects fromUnit):

| Function | Supported From Units |
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
### Recipe 15.2 — Gas Conversion

**Scenario**: Convert gas volume at specified conditions to standard state (0°C, 1 atm) volume.

```
=UDF_PC_CONVERTSTANDARD(volume, pressure, temperature, molWeight)
=UDF_PC_STDVOLUME(cylVol, fillP, fillT, gasName)
```

#### ConvertStandard

Convert gas volume at arbitrary (P, T) conditions to volume and mass at standard state (0°C, 1 atm). Returns a 2D array with V_std and mass.

**VBA Usage**
```vb
ConvertStandard(v, p, T, MW) As Variant(,)
```

| Parameter | Type | Description |
|------|------|------|
| v | Double | Volume (L) |
| p | Double | Pressure (kPa) |
| T | Double | Temperature (°C) |
| MW | Double | Molar mass (g/mol) |

```vb
Dim result As Variant
result = ConvertStandard(10, 202.65, 25, 28.01)
' result(0,0)="Std Volume(L)", result(0,1)="Std Mass(g)"
' result(1,0)=18.61, result(1,1)=23.27
```

**UDF Usage**
```
=UDF_PC_CONVERTSTANDARD(v, p, T, MW)
```

#### IdealGasLaw

PV = nRT solver. Pass three of the four parameters; pass Empty (or omit) the parameter to solve for. R = 8.314 J/(mol·K). **P in Pa, V in m³, T in K**.

```vb
IdealGasLaw(P, V, n, T) As Double
```

```vb
' Standard state 1 mol ideal gas: P=101325 Pa, V≈0.022414 m³, T=273.15 K
n = IdealGasLaw(101325, 0.022414, Empty, 273.15)  ' → ~1 (mol)

' Solve for moles: P=202650 Pa, V=0.01 m³, T=298 K → n≈0.818 mol
n = IdealGasLaw(202650, 0.01, Empty, 298)  ' → ~0.818
```

**UDF Usage**
```
=UDF_PC_IDEALGASLAW(P, V, n, T)
```

#### CompressFactorPR

Calculates the Peng-Robinson equation of state compressibility factor Z. Requires critical temperature Tc (K), critical pressure Pc (kPa), and acentric factor omega.

```vb
CompressFactorPR(P, T, Tc, Pc, omega) As Double
```

```vb
Z = CompressFactorPR(1013.25, 300, 304.2, 7380, 0.225)  ' CO2 at 1 atm, 27°C
```

**UDF Usage**
```
=UDF_PC_COMPRESS(P, T, Tc, Pc, omega)
=UDF_PC_ZFACTOR(P, T, gasName)
```

<a id="recipe-unit-conversion"></a>
### Recipe 15.3 — Unit Conversion

**Scenario**: Lab and engineering calculations frequently involve mass, volume, pressure, and temperature unit conversions. Use unified conversion functions for consistency and traceability.

**Input**
|   | A | B | C |
|---|---|---|---|
| 1 | Value | From Unit | To Unit |
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

CylinderStdVolume calculates standard state volume from cylinder volume, fill pressure, and temperature (automatically includes Z factor correction). CylinderStdVolumeFromMass calculates standard state volume from cylinder net weight and gas chemical formula.

```vb
CylinderStdVolume(cylVol, fillP, fillT, gasName) As Double
CylinderStdVolumeFromMass(netWt, gasFormula) As Double
```

| Parameter | Type | Description |
|------|------|------|
| cylVol | Double | cylinder volume (L) |
| fillP | Double | fill pressure (kPa) |
| fillT | Double | Fill temperature (°C) |
| gasName | String | Gas name (e.g. "CO2", "N2", "O2") |
| netWt | Double | cylinder net weight (kg) |
| gasFormula | String | Chemical formula (e.g. "CO2", "N2") |

```vb
V_std = CylinderStdVolume(40, 15000, 25, "N2")          ' 40L N2 Cylinder standard state volume
V_std = CylinderStdVolumeFromMass(25, "CO2")            ' 25kg CO2 Cylinder standard state volume
```

**UDF Usage**
```
=UDF_PC_STDVOLUME(cylVol, fillP, fillT, gasName)
=UDF_PC_STDVOLMASS(netWt, gasFormula)
```

---

