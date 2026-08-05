# Excel VBA Libraries

High-performance VBA function library for Excel — array operations, statistical analysis, matrix computation, string processing, and data structure transformation.

**Version 2.1.0** | ✅ All Tests Passed | All Modules | [中文](README.md)

## When to Use

When Excel's built-in functions fall short or pure VBA implementation is too complex:

- **Statistics** — descriptive stats, t-tests, correlation matrices, distribution functions — no need to leave Excel for Python/R
- **Matrix Operations** — SVD, pseudoinverse, QR, Cholesky, linear system solving — directly in cells
- **JSON Processing** — parse API responses, extract nested fields, import config tables — pure VBA, no external parser
- **XML Parsing** — XPath queries, attribute extraction, XML to worksheet — based on Windows built-in MSXML2
- **SQL Queries** — treat worksheets as databases, SELECT / JOIN / GROUP BY in a single formula
- **String Encoding** — Base64, URL encoding, HTML entities, Levenshtein edit distance
- **Data Cleaning** — regex replacement, array dedup/sort/filter, pivot/unpivot, set operations
- **Date Calculations** — ISO week numbers, workdays, age, Unix timestamp conversion
- **File Batch Processing** — UTF-8 read/write, folder traversal, batch merge
- **PhysChem** — molecular weight from formula, unit conversion (volume/pressure/temperature), ideal gas standard state
- **Report Export** — one-click export worksheet ranges to HTML / JSON / Markdown

## Key Advantages

- **Zero Dependencies** — import and use, no external components. All dependencies are Windows built-in: `Scripting.Dictionary`, `VBScript.RegExp`, `ADODB`, `MSXML2`. **Note**: 64-bit Office requires [Microsoft Access Database Engine 2016 Redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920) for SQL query features (`SqlUtils`).
- **Dual Mode** — worksheet UDFs (`=UDF_*`) and VBA code, covering both cell formulas and macro programming.
- **Numerically Stable** — Kahan compensated summation, QR decomposition instead of normal equations, two-pass variance to avoid floating-point accumulation errors.
- **Fully Tested** — module unit tests + cross-validation + comprehensive handbook examples, 100% pass rate.

## Quick Start

**Requirements**: Excel 2010+ (32/64-bit), Windows. Excel 2021 or Microsoft 365 recommended (dynamic array support, automatic formula spilling, no `Ctrl+Shift+Enter` needed). **64-bit Office**: Access Database Engine required for SqlUtils (see Key Advantages note above).

1. Press `Alt+F11` in Excel to open the VBA editor
2. Menu → **File** → **Import File** → select `.bas` modules from `src/`
3. If using `VBA-Core/` class modules, import in order: `VariantKit → ArrayOps → DictProxy`
4. Worksheet UDFs (`UDF_*` prefix) can be used directly in cell formulas
5. Functions returning `Dictionary` are VBA-only (see module header comments)
6. Save as `.xlsm` (Macro-Enabled Workbook)

> See [AGENTS.md](AGENTS.md) for module dependencies. Pre-packaged files are also available: `docs/VBA_Libraries.xlsm` (import on demand) or `docs/VBA_Libraries.xlam` (add-in, globally available), both including VBA-Core.

## How to Use

All functions support two calling modes:

**Worksheet UDF** — enter formulas directly in cells:

|   | A |
|---|---|
| 1 | 3 |
| 2 | 1 |
| 3 | 4 |
| 4 | 1 |
| 5 | 5 |

`=UDF_ARR_SORT(A1:A5, TRUE)` → `{1, 1, 3, 4, 5}`

**VBA Code** — call from the VBA editor:

```vb
Dim result As Variant
result = ArraySort(Array(3, 1, 4, 1, 5), True)  ' → Array(1, 1, 3, 4, 5)
```

> See the [User Manual (EN)](docs/VBA_LIB_User_Manual_EN.md) for more examples and full parameter descriptions, and the [API Reference](rules/api-reference.md) for function signature lookup.

## Module Overview

15 modules across 6 layers. See [AGENTS.md](AGENTS.md) for the full structure with dependencies.

| Layer | Modules |
|------|------|
| Data | ArrayUtils, DictSetUtils, PivotUtils, SqlUtils |
| Statistics/Math | LinearUtils, StatsUtils, RegressUtils |
| Text | StringUtils, RegexUtils, JsonUtils, XmlUtils |
| Date | DateTimeUtils |
| Excel/File | RangeUtils, FileSystemUtils |
| PhysChem | PhyChemUtils |

## Formula Quick Reference

| Formula | Description |
|------|------|
| `=UDF_ARR_SORT(A1:A50, FALSE)` | Descending sort |
| `=UDF_LINALG_SVD_SVALS(A1:C10)` | Matrix singular values |
| `=UDF_STAT_MEAN(A1:A100)` | Arithmetic mean |
| `=UDF_REGRESS_IMPORTANCE(A1:E20, A1:D1, 5)` | Factor importance ranking |
| `=UDF_REGEX_EXTRACT(A1, "\d{4}-\d{2}-\d{2}")` | Regex extract date |
| `=UDF_STR_LEFTOF(A1, "@")` | Text left of @ |
| `=UDF_DT_ISOWEEKNUM(TODAY())` | Today's ISO week number |
| `=UDF_JSON_GET(A1, "users[0].name")` | JSON path extraction |
| `=UDF_JSON_STRINGIFY(A1)` | Variant → JSON string |
| `=UDF_XML_GET(A1, "/a/b")` | XML XPath value |
| `=UDF_XML_TABLE(A1, "/r/i", {"a","b"})` | XML to table |
| `=UDF_PC_MOLWEIGHT("H2O")` | Molecular weight from formula |
| `=UDF_SQL_QUERY("SELECT * FROM [Sheet1$]")` | SQL query on worksheet |

See the [API Reference](rules/api-reference.md) for all function signatures and the [User Manual (EN)](docs/VBA_LIB_User_Manual_EN.md) for complete usage.

## Documentation Map

> **Information is defined in exactly one place; all others link to it.** Every document has one primary audience.

| Document | Audience | When to Read |
|------|------|---------|
| [User Manual (EN)](docs/VBA_LIB_User_Manual_EN.md) | End users | Recipe-driven, by scenario |
| [User Manual (CN)](rules/user-manual.md) | Chinese users | Chinese edition |
| [API Reference](rules/api-reference.md) | Developers | Function signature lookup |
| [AGENTS.md](AGENTS.md) | Developers / AI | Project structure, dependencies, coding conventions |
| [README (CN)](README.md) | Chinese readers | Project homepage (Chinese) |

> The complete documentation system (including VBA coding standards, Python test guide, SQL conventions, manual authoring rules) is detailed in the [AGENTS.md](AGENTS.md) document routing table.

## Testing

> ✅ All tests passed: module unit tests + cross-validation + comprehensive handbook examples.
> See [AGENTS.md](AGENTS.md) for test framework and run commands. Requires Python 3 + pywin32 + numpy + scipy + Excel (Windows).

<!-- last_updated: 2026-06-17 -->
