# Excel VBA Libraries -- API Reference

<!-- last_updated: 2026-08-01 -->

> Function signature quick reference. For full usage and examples see [User Manual (CN)](VBA_LIB_User_Manual.md) / [User Manual (EN)](VBA_LIB_User_Manual_EN.md).

<!-- AUTO_COUNTS_START -->
**15 模块 | 533 Public Functions | 33 Public Subs | 共 566 个 Public 接口**
<!-- AUTO_COUNTS_END -->

---

## ArrayUtils -- Array Operations

**Module**: `ArrayUtils.bas` | **Public functions**: 32 | **UDFs**: 30

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| ArrayDims | `(arr)` | Long | -- |
| IsArray1D | `(arr)` | Boolean | -- |
| ArrayUnique | `(arr)` | Variant | UDF_ARR_UNIQUE |
| ArraySort | `(arr, [ascending])` | Variant | UDF_ARR_SORT |
| ArrayFilterByValue | `(arr, matchValue, [operator])` | Variant | UDF_ARR_FILTER |
| ArrayCountIf | `(arr, matchValue, [operator])` | Long | UDF_ARR_COUNTIF |
| ArrayConcat | `(arr1, arr2)` | Variant | UDF_ARR_CONCAT |
| ArraySlice | `(arr, [start], [cnt])` | Variant | UDF_ARR_SLICE |
| ArrayFlatten | `(arr)` | Variant | UDF_ARR_FLATTEN |
| ArrayTranspose1D | `(arr, [asColumn])` | Variant | UDF_ARR_TRANSPOSE |
| ArrayFind | `(arr, value, [caseSensitive])` | Variant | UDF_ARR_FIND |
| ArrayContains | `(arr, value, [caseSensitive])` | Variant | UDF_ARR_CONTAINS |
| ArrayLookup | `(lookupArray, lookupValue, lookupCol, [returnCols], [matchType])` | Variant | UDF_ARR_LOOKUP |
| ArrayShuffle | `(arr)` | Variant | UDF_ARR_SHUFFLE |
| ArraySample | `(arr, n, [withReplacement])` | Variant | UDF_ARR_SAMPLE |
| LinSpace | `(start, endVal, n)` | Variant | UDF_ARR_LINSPACE |
| RangeFill | `(start, count, [stepSize])` | Variant | UDF_ARR_RANGEFILL |
| ArrayChunk | `(arr, size)` | Variant | UDF_ARR_CHUNK |
| ArrayMin | `(arr)` | Variant | UDF_ARR_MIN |
| ArrayMax | `(arr)` | Variant | UDF_ARR_MAX |
| ArraySum | `(arr)` | Variant | UDF_ARR_SUM |
| ArrayToString | `(arr, [delimiter])` | Variant | UDF_ARR_TOSTRING |
| ArrayReverse | `(arr)` | Variant | UDF_ARR_REVERSE |
| ArrayGetRow | `(arr, row)` | Variant | UDF_ARR_GETROW |
| ArrayGetCol | `(arr, col)` | Variant | UDF_ARR_GETCOL |
| ArrayTranspose2D | `(arr)` | Variant | UDF_ARR_TRANSPOSE2D |
| ArrayEqual | `(arr1, arr2, [caseSensitive])` | Variant | UDF_ARR_EQUAL |
| ArrayProduct | `(arr)` | Variant | UDF_ARR_PRODUCT |
| CumSum | `(arr)` | Variant | UDF_ARR_CUMSUM |
| ArgSort | `(arr, [ascending])` | Variant | UDF_ARR_ARGSORT |
| ArrayAny | `(arr, matchValue, [operator])` | Variant | UDF_ARR_ANY |
| ArrayAll | `(arr, matchValue, [operator])` | Variant | UDF_ARR_ALL |

---

## DictSetUtils -- Dictionary and Set Operations

**Module**: `DictSetUtils.bas` | **Public functions**: 28 | **UDFs**: 8

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| DictMerge | `(dict1, dict2, [overwrite])` | Object | -- |
| DictMergeSum | `(dict1, dict2)` | Object | -- |
| DictInvert | `(dict)` | Object | -- |
| DictKeys | `(dict)` | Variant | -- |
| DictValues | `(dict)` | Variant | -- |
| DictTo2DArray | `(dict)` | Variant | -- |
| ArrayToDict | `(arr, [keyFn])` | Object | -- |
| DictFrom2DArray | `(arr, [colKey], [colValue])` | Object | -- |
| CountFrequency | `(arr)` | Object | -- |
| GroupCount | `(arr)` | Variant | UDF_DICT_GROUPCOUNT |
| DictPick | `(dict, keys)` | Object | -- |
| DictCount | `(dict)` | Long | -- |
| SetUnion | `(arr1, arr2)` | Variant | UDF_DICT_UNION |
| SetIntersect | `(arr1, arr2)` | Variant | UDF_DICT_INTERSECT |
| SetDifference | `(arr1, arr2)` | Variant | UDF_DICT_DIFFERENCE |
| SetSymDifference | `(arr1, arr2)` | Variant | UDF_DICT_SYM_DIFF |
| SetIsSubset | `(arr1, arr2)` | Boolean | UDF_DICT_ISSUBSET |
| SetEqual | `(arr1, arr2)` | Boolean | UDF_DICT_ISEQUAL |
| DictFilterByValue | `(dict, matchValue, [operator])` | Object | -- |
| DictSortByKey | `(dict, [ascending])` | Variant | -- |
| DictSortByValue | `(dict, [ascending])` | Variant | -- |
| DictGetDefault | `(dict, key, [defaultValue])` | Variant | -- |
| DictRenameKey | `(dict, oldKey, newKey)` | Object | -- |
| DictRemoveKeys | `(dict, keys)` | Object | -- |
| DictClone | `(dict)` | Object | -- |
| DictIsEmpty | `(dict)` | Boolean | -- |
| DictTopN | `(dict, n, [ascending])` | Variant | -- |
| SetCartesianProduct | `(arrA, arrB)` | Variant | UDF_DICT_CARTESIAN |

---

## PivotUtils -- Data Reshaping

**Module**: `PivotUtils.bas` | **Public functions**: 5 | **Public subs**: 6 | **UDFs**: 4

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| RawConversion | `(srcRange, valueColRef, colDimRef, rowDimRef, [keepBlank], [sortLabels])` | Variant | UDF_PIVOT_CONVERT |
| RawConversionFromArray | `(dataArray, valueColName, colDimName, rowDimName, [keepBlank], [sortLabels])` | Variant | -- |
| RawConversionToRange (Sub) | `(srcRange, valueColRef, colDimRef, rowDimRef, destCell, [keepBlank], [sortLabels])` | -- | -- |
| Unpivot (Sub) | `(rng, valueCols, nameCol, valCol, destCell, [idColIndices])` | -- | -- |
| GroupBy | `(rng, groupCol, aggCol, [aggFunc], [destCell])` | Variant | UDF_PIVOT_GROUPBY |
| SplitColumnToRows (Sub) | `(rng, colIdx, delimiter, destCell)` | -- | -- |
| MergeColumns (Sub) | `(rng, colIndices, delimiter, newColName, destCell, [dropOrigCols])` | -- | -- |
| FilterTable (Sub) | `(rng, colIdx, op, value, destCell)` | -- | -- |
| TransposeTable (Sub) | `(rng, destCell)` | -- | -- |
| VLookupArray | `(dataArray, lookupValue, lookupCol, [returnCol])` | Variant | UDF_PIVOT_VLOOKUP |
| CrossJoin | `(rng1, rng2)` | Variant | UDF_PIVOT_CROSSJOIN |

---

## SqlUtils -- SQL Queries

**Module**: `SqlUtils.bas` | **Public functions**: 11 | **Public subs**: 1 | **UDFs**: 6

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| SqlEscapeString | `(value)` | String | -- |
| MakeSafeColumnName | `(colName)` | String | -- |
| CloseSqlCache (Sub) | `()` | -- | -- |
| SqlGetConnection | `([filePath], [outOk])` | Object | -- |
| SqlExecute | `(sql, [filePath], [includeHeader], outOk, [outErrorMsg])` | Variant() | -- |
| SqlQuery | `(selectClause, fromClause, [whereClause], [orderByClause], [filePath], [outOk])` | Variant() | UDF_SQL_QUERY |
| SqlJoin | `(table1, table2, joinOn, [joinType], [selectCols], [filePath], [outOk])` | Variant() | UDF_SQL_JOIN |
| SqlGroupBy | `(tableName, groupCols, aggExprs, [whereClause], [filePath], [outOk])` | Variant() | UDF_SQL_GROUPBY |
| SqlListSheets | `([filePath], [outOk])` | Variant() | UDF_SQL_LIST_SHEETS |
| SqlListColumns | `(tableName, [filePath], [outOk])` | Variant() | UDF_SQL_LIST_COLUMNS |
| SqlListTables | `([filePath], [outOk])` | Variant() | UDF_SQL_LIST_TABLES |
| SqlRangeQuery | `(sql, rng, [tableAlias], [outOk])` | Variant() | -- |

---

## LinearUtils -- Linear Algebra

**Module**: `LinearUtils.bas` | **Public functions**: 27 | **Public subs**: 9 | **UDFs**: 14

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| MatrixRows | `(A)` | Long | -- |
| MatrixCols | `(A)` | Long | -- |
| MatrixFrobeniusNorm | `(A)` | Double | -- |
| RangeToMatrix | `(rng)` | Double() | -- |
| MatrixToRange (Sub) | `(mat, startCell)` | -- | -- |
| SelectionToArray2D | `(rng)` | Variant() | -- |
| ArrayToRange (Sub) | `(data, startCell, [asColumn])` | -- | -- |
| MatrixTranspose | `(A)` | Double() | -- |
| IdentityMatrix | `(n)` | Double() | -- |
| MatrixCopy | `(A)` | Double() | -- |
| MatrixGetColumn | `(A, k)` | Double() | -- |
| MatrixSetColumn (Sub) | `(A, k, col)` | -- | -- |
| MatrixMultiply | `(A, B, [blockSize])` | Double() | -- |
| MatrixMultiplyNaive | `(A, B)` | Double() | -- |
| SVD (Sub) | `(A, U, S, Vt, [tol], [maxSweeps])` | -- | UDF_LINALG_SVD_U / S / VT / SVALS |
| PseudoInverse | `(A, [tolerance])` | Double() | UDF_LINALG_PINV |
| MatrixRank_Array | `(A, [tolerance])` | Long | UDF_LINALG_RANK |
| EigenSymmetric (Sub) | `(A, V, D, [tol], [maxSweeps])` | -- | UDF_LINALG_EIGVAL / EIGVEC |
| QRDecomposition (Sub) | `(A, Q, R, [economy])` | -- | UDF_LINALG_QR_Q / QR_R |
| QRDecompositionPiv (Sub) | `(A, Q, R, perm, [economy])` | -- | -- |
| MatrixTrace | `(A)` | Double | -- |
| MatrixScale | `(A, scalar)` | Double() | -- |
| MatrixAdd | `(A, B)` | Double() | -- |
| MatrixSubtract | `(A, B)` | Double() | -- |
| MatrixHadamard | `(A, B)` | Double() | -- |
| MatrixNorm | `(A, [normType])` | Double | -- |
| MatrixPower | `(A, n)` | Double() | -- |
| LUDecomposition (Sub) | `(A, L, U, P, [outSwapCount])` | -- | -- |
| CholeskyDecomposition (Sub) | `(A, L)` | -- | UDF_LINALG_CHOLESKY |
| MatrixDeterminant | `(A)` | Double | UDF_LINALG_DET |
| SolveLinearSystem | `(A, b, [tolerance])` | Double() | UDF_LINALG_SOLVE |
| MatrixConditionNumber | `(A, [tol], [maxSweeps])` | Double | -- |
| VectorDot | `(a, b)` | Double | -- |
| VectorNorm | `(v, [normType])` | Double | -- |
| VectorCross | `(a, b)` | Double() | -- |
| PolyFit | `(rngX, rngY, [degree])` | Variant | UDF_LINALG_POLYFIT |

---

## StatsUtils -- Statistics

**Module**: `StatsUtils.bas` | **Public functions**: 44 | **UDFs**: 33

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| Mean | `(data, [colIndex])` | Variant | UDF_STAT_MEAN |
| WeightedMean | `(values, weights, [colIdxV], [colIdxW])` | Variant | UDF_STAT_WMEAN |
| Median | `(data, [colIndex])` | Variant | UDF_STAT_MEDIAN |
| Mode | `(data, [colIndex])` | Variant | UDF_STAT_MODE |
| Min | `(data, [colIndex])` | Variant | -- |
| Max | `(data, [colIndex])` | Variant | -- |
| MinMax | `(data, [outMin], [outMax], [colIndex])` | Variant | UDF_STAT_MINMAX |
| GeometricMean | `(data, [colIndex])` | Variant | UDF_STAT_GEOMEAN |
| HarmonicMean | `(data, [colIndex])` | Variant | UDF_STAT_HARMEAN |
| TrimMean | `(data, [trimPct], [colIndex])` | Variant | UDF_STAT_TRIMEAN |
| RootMeanSquare | `(data, [colIndex])` | Variant | UDF_STAT_RMS |
| MeanAbsDev | `(data, [colIndex])` | Variant | UDF_STAT_MAD |
| StdDev | `(data, [colIndex])` | Variant | UDF_STAT_STDEV |
| StdDevP | `(data, [colIndex])` | Variant | UDF_STAT_STDEVP |
| Variance | `(data, [colIndex])` | Variant | UDF_STAT_VAR |
| VarianceP | `(data, [colIndex])` | Variant | UDF_STAT_VARP |
| Percentile | `(data, k, [colIndex])` | Variant | UDF_STAT_PERCENTILE |
| IQR | `(data, [colIndex])` | Variant | UDF_STAT_IQR |
| Skewness | `(data, [colIndex])` | Variant | UDF_STAT_SKEW |
| Kurtosis | `(data, [colIndex])` | Variant | UDF_STAT_KURTOSIS |
| Rank | `(data, value, [ascending], [colIndex])` | Variant | UDF_STAT_RANK |
| RankEq | `(data, value, [ascending], [colIndex])` | Variant | UDF_STAT_RANKEQ |
| RankAvg | `(data, value, [ascending], [colIndex])` | Variant | UDF_STAT_RANKAVG |
| PercentRank | `(data, value, [ascending], [colIndex])` | Variant | UDF_STAT_PERCENTRANK |
| Covariance | `(dataX, dataY, [colX], [colY])` | Variant | UDF_STAT_COV |
| Correlation | `(dataX, dataY, [colX], [colY])` | Variant | UDF_STAT_CORREL |
| RSquare | `(actual, predicted, [colActual], [colPredicted])` | Variant | UDF_STAT_R2 |
| CorrelationMatrix | `(data, [hasHeader])` | Variant() | -- |
| ZScore | `(data, [value], [colIndex])` | Variant | UDF_STAT_ZSCORE |
| Normalize | `(data, [colIndex])` | Variant | UDF_STAT_NORMALIZE |
| LinInterp | `(x, xs, ys)` | Variant | UDF_STAT_LININTERP |
| Winsorize | `(data, [pct], [colIndex])` | Variant | UDF_STAT_WINSORIZE |
| MovingAverage | `(data, window, [colIndex])` | Variant | UDF_STAT_MA |
| Binning | `(data, nBins, [colIndex])` | Variant | -- |
| ZTest | `(data, mu0, sigma, [colIndex])` | Variant | UDF_STAT_ZTEST |
| TTest | `(data1, data2, [testType], [colIdx1], [colIdx2])` | Variant | UDF_STAT_TTEST |
| StandardError | `(data, [colIndex])` | Variant | UDF_STAT_SE |
| ConfidenceInterval | `(data, [alpha], [colIndex])` | Variant | -- |
| GammaLn | `(x)` | Double | -- |
| BetaReg | `(x, a, b)` | Double | -- |
| TDistCDF | `(tVal, df)` | Double | -- |
| TDist2T | `(tStat, df)` | Double | -- |
| FDistRT | `(fStat, df1, df2)` | Double | -- |
| TInv2T | `(alpha, df)` | Double | -- |

---

## RegressUtils -- Regression and Factor Analysis

**Module**: `RegressUtils.bas` | **Public functions**: 9 | **UDFs**: 7

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| FactorImportance | `(data, factorCols, resultCol, [hasHeader])` | Variant() | UDF_REGRESS_IMPORTANCE |
| InteractionEffects | `(data, factorCols, resultCol, [hasHeader])` | Variant() | UDF_REGRESS_INTERACT |
| ANOVAOneWay | `(data, factorCol, resultCol, [hasHeader])` | Object | UDF_REGRESS_ANOVA |
| ANOVAOneWay_Fstat | `(data, factorCol, resultCol, [hasHeader])` | Double | -- |
| FitOLS | `(X, y)` | Object | -- |
| LinearModelFit | `(data, factorCols, resultCol, [hasHeader])` | Object | -- |
| LinearModelPredict | `(model, newData)` | Double | UDF_REGRESS_PREDICT |
| FactorSweep | `(model, factorIndex, fromVal, toVal, steps, [baseValues])` | Variant() | UDF_REGRESS_SWEEP |
| OptimizeFactors | `(data, factorCols, resultCol, [goal], [topN], [nSteps], [hasHeader])` | Variant() | UDF_REGRESS_OPTIMIZE |
| CorrelationMatrix↗ | `(data, [hasHeader])` | Variant() | UDF_REGRESS_CORREL |

> Note: `ANOVAOneWay_Fstat` returns the F-statistic directly. `UDF_REGRESS_CORREL` delegates to `CorrelationMatrix` in StatsUtils.bas.

---

## StringUtils -- String Processing

**Module**: `StringUtils.bas` | **Public functions**: 35 | **UDFs**: 35

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| ExtractBetween | `(text, leftDelim, rightDelim, [nth], [includeDelim])` | String | UDF_STR_EXTRACTBETWEEN |
| RemoveDiacritics | `(text)` | String | UDF_STR_REMOVEDIACRITICS |
| ReverseString | `(text)` | String | UDF_STR_REVERSESTRING |
| CountSubstring | `(text, search, [caseSensitive])` | Long | UDF_STR_COUNTSUBSTRING |
| StartsWith | `(text, prefix, [caseSensitive])` | Boolean | UDF_STR_STARTSWITH |
| EndsWith | `(text, suffix, [caseSensitive])` | Boolean | UDF_STR_ENDSWITH |
| LeftOf | `(text, delimiter, [nth])` | String | UDF_STR_LEFTOF |
| RightOf | `(text, delimiter, [nth], [fromRight])` | String | UDF_STR_RIGHTOF |
| TextJoin | `(delimiter, ...args)` | String | UDF_STR_TEXTJOIN |
| NthWord | `(text, n, [delimiter])` | String | UDF_STR_NTHWORD |
| CommonPrefix | `(a, b, [caseSensitive])` | String | UDF_STR_COMMONPREFIX |
| PadLeft | `(text, totalWidth, [padChar])` | String | UDF_STR_PADLEFT |
| PadRight | `(text, totalWidth, [padChar])` | String | UDF_STR_PADRIGHT |
| Repeat | `(text, n)` | String | UDF_STR_REPEAT |
| NormalizeWhitespace | `(text)` | String | UDF_STR_NORMALIZEWHITESPACE |
| RemoveChars | `(text, charsToRemove)` | String | UDF_STR_REMOVECHARS |
| KeepChars | `(text, allowedChars)` | String | UDF_STR_KEEPCHARS |
| ToTitleCase | `(text)` | String | UDF_STR_TOTITLECASE |
| IsNullOrEmpty | `(text)` | Boolean | UDF_STR_ISNULLOREMPTY |
| IsNullOrWhitespace | `(text)` | Boolean | UDF_STR_ISNULLORWHITESPACE |
| IsEmail | `(text)` | Boolean | UDF_STR_ISEMAIL |
| IsUrl | `(text)` | Boolean | UDF_STR_ISURL |
| Coalesce | `(...values)` | Variant | UDF_STR_COALESCE |
| Truncate | `(text, maxLength, [suffix])` | String | UDF_STR_TRUNCATE |
| UUID | `()` | String | UDF_STR_UUID |
| RandomString | `([length], [charset])` | String | UDF_STR_RANDOMSTRING |
| Slugify | `(text, [separator])` | String | UDF_STR_SLUGIFY |
| URLEncode | `(text)` | String | UDF_STR_URLENCODE |
| URLDecode | `(text)` | String | UDF_STR_URLDECODE |
| Base64Encode | `(text, [encoding])` | String | UDF_STR_BASE64ENCODE |
| Base64Decode | `(base64, [encoding])` | String | UDF_STR_BASE64DECODE |
| HTMLEncode | `(text)` | String | UDF_STR_HTMLENCODE |
| HTMLDecode | `(text)` | String | UDF_STR_HTMLDECODE |
| Soundex | `(text)` | String | UDF_STR_SOUNDEX |
| LevenshteinDistance | `(a, b, [caseSensitive])` | Long | UDF_STR_LEVENSHTEIN |

---

## RegexUtils -- Regular Expressions

**Module**: `RegexUtils.bas` | **Public functions**: 10 | **Public subs**: 2 | **UDFs**: 9

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| RegexIsMatch | `(inputText, pattern, [ignoreCase], [multiLine])` | Variant | UDF_REGEX_ISMATCH |
| RegexExtract | `(inputText, pattern, [instance], [ignoreCase], [multiLine], [separator])` | Variant | UDF_REGEX_EXTRACT |
| RegexReplace | `(inputText, pattern, replacement, [instance], [ignoreCase], [multiLine], [globalReplace])` | Variant | UDF_REGEX_REPLACE |
| RegexSplit | `(inputText, pattern, [ignoreCase], [multiLine])` | Variant | UDF_REGEX_SPLIT |
| RegexCount | `(inputText, pattern, [ignoreCase], [multiLine])` | Variant | UDF_REGEX_COUNT |
| RegexSplitToRange (Sub) | `(srcRange, pattern, destCell, [ignoreCase], [multiLine], [fillValue])` | -- | -- |
| RegexExtractAll | `(inputText, pattern, [ignoreCase], [multiLine])` | Variant | UDF_REGEX_EXTRACTALL |
| RegexExtractGroups | `(inputText, pattern, [ignoreCase], [multiLine])` | Variant | UDF_REGEX_EXTRACTGROUPS |
| RegexIsFullMatch | `(inputText, pattern, [ignoreCase], [multiLine])` | Variant | UDF_REGEX_FULLMATCH |
| RegexEscape | `(text)` | Variant | UDF_REGEX_ESCAPE |
| RegexFindInRange | `(rng, pattern, [ignoreCase], [multiLine])` | Range | -- |
| RegexReplaceRange (Sub) | `(rng, pattern, replacement, [ignoreCase], [multiLine])` | -- | -- |

---

## JsonUtils -- JSON Processing

**Module**: `JsonUtils.bas` | **Public functions**: 5 | **Public subs**: 1 | **UDFs**: 4

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| JsonParse | `(json)` | Variant | -- |
| JsonGet | `(json, path)` | Variant | UDF_JSON_GET |
| JsonToRange (Sub) | `(json, startCell)` | -- | -- |
| JsonGetKeys | `(json)` | String() | UDF_JSON_KEYS |
| JsonIsValid | `(json)` | Boolean | UDF_JSON_IS_VALID |
| JsonStringify | `(value)` | String | UDF_JSON_STRINGIFY |

---

## XmlUtils -- XML to Excel

**Module**: `XmlUtils.bas` | **Public functions**: 4 | **UDFs**: 3

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| XmlValidate | `(xml, [errDetail])` | Boolean | UDF_XML_VALIDATE |
| XmlGet | `(xml, xpath)` | Variant | UDF_XML_GET |
| XmlGetAttr | `(xml, xpath, attrName)` | Variant | -- |
| XmlToRange | `(xml, rowXPath, [colNames])` | Variant | UDF_XML_TABLE |

> **Dependencies**: MSXML2.DOMDocument (Windows built-in). XPath: `/a/b`, `//c`, `[@k='v']`, `[1]`.

---


## DateTimeUtils -- Date and Time

**Module**: `DateTimeUtils.bas` | **Public functions**: 27 | **UDFs**: 25

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| ISOWeekNum | `(d)` | Long | UDF_DT_ISOWEEKNUM |
| FirstDayOfMonth | `([d])` | Date | UDF_DT_FIRSTDAYOFMONTH |
| LastDayOfMonth | `([d])` | Date | UDF_DT_LASTDAYOFMONTH |
| FirstDayOfQuarter | `([d])` | Date | UDF_DT_FIRSTDAYOFQUARTER |
| LastDayOfQuarter | `([d])` | Date | UDF_DT_LASTDAYOFQUARTER |
| DaysInMonth | `([y], [m])` | Long | UDF_DT_DAYSINMONTH |
| DaysInYear | `([y])` | Long | UDF_DT_DAYSINYEAR |
| DayOfYear | `(d)` | Long | UDF_DT_DAYOFYEAR |
| IsLeapYear | `(y)` | Boolean | UDF_DT_ISLEAPYEAR |
| Quarter | `([d])` | Long | UDF_DT_QUARTER |
| Age | `(birthDate, [asOf])` | Object | -- |
| AgeYears | `(birthDate, [asOf])` | Long | UDF_DT_AGEYEARS |
| WorkdaysBetween | `(startDate, endDate, [holidays], [outOk])` | Long | UDF_DT_WORKDAYSBETWEEN |
| NextWorkday | `(startDate, n, [holidays], [outOk])` | Date | UDF_DT_NEXTWORKDAY |
| DateRange | `(start, endDate, [interval])` | Variant | UDF_DT_DATERANGE |
| NthWeekday | `(y, m, n, [dayOfWeek])` | Variant | UDF_DT_NTHWEEKDAY |
| FiscalYear | `([d], [startMonth])` | Long | UDF_DT_FISCALYEAR |
| UnixToDate | `(ts)` | Date | UDF_DT_UNIXTODATE |
| DateToUnix | `(d)` | Double | UDF_DT_DATETOUNIX |
| IsWeekend | `(d)` | Boolean | UDF_DT_ISWEEKEND |
| IsHoliday | `(d, holidays)` | Boolean | UDF_DT_ISHOLIDAY |
| AddMonthsSafe | `(d, n)` | Date | UDF_DT_ADDMONTHSSAFE |
| DateDiffParts | `(startDate, endDate)` | Object | -- |
| WeekOfMonth | `(d, [startDay])` | Long | UDF_DT_WEEKOFMONTH |
| StartOfWeek | `(d, [startDay])` | Date | UDF_DT_STARTOFWEEK |
| EndOfWeek | `(d, [startDay])` | Date | UDF_DT_ENDOFWEEK |
| Easter | `(y)` | Date | UDF_DT_EASTER |

---

## RangeUtils -- Range Operations

**Module**: `RangeUtils.bas` | **Public functions**: 23 | **Public subs**: 10 | **UDFs**: 16

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| LastRow | `([ws])` | Long | UDF_RANGE_LASTROW |
| LastCol | `([ws])` | Long | UDF_RANGE_LASTCOL |
| FirstRow | `([ws])` | Long | UDF_RANGE_FIRSTROW |
| FirstCol | `([ws])` | Long | UDF_RANGE_FIRSTCOL |
| ColLetter | `(colNum)` | String | UDF_RANGE_COL_LETTER |
| ColNumber | `(colRef)` | Long | UDF_RANGE_COL_NUM |
| UsedRangeEx | `([ws])` | Range | -- |
| RangeToHTML | `(rng, [includeHeader], [tableClass])` | String | UDF_RANGE_TOHTML |
| SafeText | `(v)` | String | UDF_RANGE_SAFETEXT |
| FindAll | `(rng, value, [lookIn], [lookAt], [matchCase], [outTruncated])` | Range | -- |
| MergeRanges | `(...ranges)` | Range | -- |
| IntersectRanges | `(...ranges)` | Range | -- |
| RangeToJSON | `(rng, [includeHeader])` | String | UDF_RANGE_TOJSON |
| RangeToMarkdown | `(rng, [includeHeader])` | String | UDF_RANGE_TOMD |
| ExportRangeToCSV (Sub) | `(rng, filePath, [delimiter], [bom], [enc], [includeHeader])` | -- | -- |
| FilterRangeToArray | `(rng, colIdx, op, value, [includeHeader])` | Variant | UDF_RANGE_FILTER |
| ValuesEqual | `(a, b)` | Boolean | -- |
| RangeToArray | `(rng)` | Variant() | -- |
| RangeExists | `(rng)` | Boolean | UDF_RANGE_EXISTS |
| NamedRangeExists | `(rangeName, [wb])` | Boolean | UDF_RANGE_NAMEDEXISTS |
| NamedRangeAdd (Sub) | `(rangeName, rng, [wb])` | -- | -- |
| NamedRangeDelete (Sub) | `(rangeName, [wb])` | -- | -- |
| IsRangeEmpty | `(rng)` | Boolean | UDF_RANGE_ISEMPTY |
| CountVisible | `(rng)` | Long | UDF_RANGE_COUNTVISIBLE |
| RangeDiff | `(rng1, rng2, [outTruncated])` | Range | -- |
| GetCellAddress | `(row, col, [absolute])` | String | UDF_RANGE_CELLADDRESS |
| CopyRangeToSheet (Sub) | `(data, startCell)` | -- | -- |
| TransposeRange (Sub) | `(rng, destCell)` | -- | -- |
| ClearRangeContents (Sub) | `(rng)` | -- | -- |
| ClearRangeFormats (Sub) | `(rng)` | -- | -- |
| AutoFitRange (Sub) | `(rng, [fitCols], [fitRows])` | -- | -- |
| RemoveDuplicatesRange (Sub) | `(rng, [colIndices], [hasHeader])` | -- | -- |
| SortRange (Sub) | `(rng, keyCol, [ascending], [hasHeader])` | -- | -- |

---

## FileSystemUtils -- File System Operations

**Module**: `FileSystemUtils.bas` | **Public functions**: 25 | **Public subs**: 3 | **UDFs**: 23

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| NormalizePath | `(path)` | String | UDF_FS_NORMALIZEPATH |
| PathCombine | `(folderPath, fileName)` | String | UDF_FS_PATHCOMBINE |
| FileExists | `(path)` | Boolean | UDF_FS_FILEEXISTS |
| FolderExists | `(path)` | Boolean | UDF_FS_FOLDEREXISTS |
| ReadTextFile | `(filePath, [encoding], [lines])` | String | UDF_FS_READTEXT |
| WriteTextFile (Sub) | `(filePath, content, [encoding], [append], [bom])` | -- | -- |
| ListFiles | `(folder, [pattern], [recurse])` | Variant | UDF_FS_LISTFILES |
| ListFolders | `(folder, [recurse])` | Variant | UDF_FS_LISTFOLDERS |
| GetFileName | `(path)` | String | UDF_FS_FILENAME |
| GetBaseName | `(path)` | String | UDF_FS_BASENAME |
| GetExtension | `(path)` | String | UDF_FS_EXTENSION |
| GetFolderPath | `(path)` | String | UDF_FS_FOLDERPATH |
| GetFileSize | `(filePath)` | Double | UDF_FS_FILESIZE |
| GetFileSizeFmt | `(filePath)` | String | UDF_FS_FILESIZEFMT |
| FileModified | `(filePath)` | Date | UDF_FS_FILEMODIFIED |
| EnsureFolder | `(folderPath)` | Boolean | UDF_FS_ENSUREFOLDER |
| TempFileName | `([prefix], [ext], [folder])` | String | UDF_FS_TEMPFILENAME |
| DeleteFile | `(filePath)` | Boolean | UDF_FS_DELETEFILE |
| CopyFileSafe | `(sourcePath, destPath, [overwrite])` | Boolean | UDF_FS_COPYFILE |
| AppendTextFile (Sub) | `(filePath, content, [encoding])` | -- | -- |
| ReadBinaryFile | `(filePath, [outOk])` | Byte() | -- |
| WriteBinaryFile (Sub) | `(filePath, bytes)` | -- | -- |
| GetTempFolder | `()` | String | UDF_FS_TEMPFOLDER |
| GetSpecialFolder | `(folderType)` | String | UDF_FS_SPECIALFOLDER |
| IsPathValid | `(path)` | Boolean | UDF_FS_ISPATHVALID |
| GetDriveInfo | `(driveLetter)` | Object | -- |
| CopyFolder | `(sourcePath, destPath, [overwrite])` | Boolean | UDF_FS_COPYFOLDER |
| DeleteFolder | `(folderPath)` | Boolean | UDF_FS_DELETEFOLDER |

---

## PhyChemUtils -- Physical Chemistry

**Module**: `PhyChemUtils.bas` | **Public functions**: 15 | **UDFs**: 16

| Function | Signature | Returns | UDF |
|----------|-----------|---------|-----|
| MolecularWeight | `(formula)` | Variant | UDF_PC_MOLWEIGHT |
| ConvertVolume | `(initialVolume, fromUnit, [toUnit])` | Variant | UDF_PC_CONVERTVOLUME |
| ConvertPressure | `(initialPressure, fromUnit, [toUnit])` | Variant | UDF_PC_CONVERTPRESSURE |
| ConvertTemperature | `(initialTemperature, fromUnit, [toUnit])` | Variant | UDF_PC_CONVERTTEMPERATURE |
| ConvertStandard | `(volume_m3, pressure_Pa, temperature_K, molWeight)` | Variant | UDF_PC_CONVERTSTANDARD |
| MassToMoles | `(mass, molWeight)` | Variant | UDF_PC_MASSTOMOLES |
| MolesToMass | `(moles, molWeight)` | Variant | UDF_PC_MOLESTOMASS |
| DilutionSolve | `(c1, v1, [c2], [v2])` | Variant | UDF_PC_DILUTION |
| IdealGasLaw | `(pressure_Pa, volume_m3, moles, temperature_K)` | Variant | UDF_PC_IDEALGASLAW |
| ConvertMass | `(initialMass, fromUnit, [toUnit])` | Variant | UDF_PC_CONVERTMASS |
| Density | `(m, V, [rho])` | Variant | UDF_PC_DENSITY |
| PercentYield | `(actual, theoretical)` | Variant | UDF_PC_YIELD |
| CompressFactorPR | `(pressure_Pa, temperature_K, Tc_K, Pc_Pa, omega)` | Variant | UDF_PC_COMPRESS, UDF_PC_ZFACTOR |
| CylinderStdVolume | `(cylVol, fillP, fillT, gasName)` | Variant | UDF_PC_STDVOLUME |
| CylinderStdVolumeFromMass | `(netWeight, gasFormula)` | Variant | UDF_PC_STDVOLMASS |

> Note: `UDF_PC_COMPRESS` hardcodes ω=0 for nonpolar gases (N₂, O₂, Ar). Use `UDF_PC_ZFACTOR` for accurate gas-specific compressibility via table lookup.

---

---
## Appendix: VBA-Core Infrastructure

Core utility classes imported by the library modules above. Import order: **VariantKit > ArrayOps > DictProxy**.

### VariantKit -- Type Normalization and Introspection

**Module**: `VariantKit.cls` | **Public functions**: 16 | **Public subs**: 1

| Function | Signature | Returns |
|----------|-----------|---------|
| IsEmptyArray | `(v)` | Boolean |
| ArrayDims | `(v)` | Long |
| Is1D | `(v)` | Boolean |
| Is2D | `(v)` | Boolean |
| IsNumericArray | `(v)` | Boolean |
| IsNumericCell | `(v)` | Boolean |
| Normalize1D | `(v, [order])` | Variant |
| NormalizeTo2D | `(data, numRows, numCols)` | Variant() |
| Normalize2D | `(v)` | Variant |
| NormalizeInput | `(v, [flattenColumn])` | Boolean |
| ToDoubles | `(v)` | Double() |
| WrapScalar | `(v)` | Variant() |
| SafeKey | `(v)` | String |
| ValuesEqual | `(a, b)` | Boolean |
| Compare | `(a, b)` | Long |
| VarLetSet (Sub) ⚠️ | `(target, source)` | -- |
| FilterPasses | `(element, matchValue, op)` | Boolean |

### ArrayOps -- Common Array Operations

**Module**: `ArrayOps.cls` | **Public functions**: 4 | **Public subs**: 2

| Function | Signature | Returns |
|----------|-----------|---------|
| Sort (Sub) | `(arr, [ascending], [comparer])` | -- |
| Slice | `(arr, start, [length])` | Variant |
| IndexOf | `(arr, value)` | Long |
| Flatten | `(arr, [order])` | Variant |
| SortIndices (Sub) | `(values, indices, [asc], [comparer])` | -- |
| CollectNumericColumns | `(dataArr, numRows, numCols, colNames, [hasHeader], [vk])` | Long() |

### DictProxy -- Safe Dictionary and Batch Operations

**Module**: `DictProxy.cls` | **Public functions**: 4

| Function | Signature | Returns |
|----------|-----------|---------|
| Create | `([compareMode])` | Object |
| FromKeys | `(keys, [defaultValue], [compareMode])` | Object |
| ToArray | `(dict)` | Variant |
| Merge | `(dict1, dict2, [overwrite])` | Object |
